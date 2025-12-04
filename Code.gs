// --- GLOBAALIT MUUTTUJAT ----
var SETTINGS_SHEET_NAME = 'Settings';
var TRANSLATIONS_SHEET_NAME = 'Translations';
var CATEGORY_SETTING_KEY = 'ExpenseCategories'; 

// --- PÄÄTOIMINNOT (KÄYTTÖLIITTYMÄ) ---

/**
 * Palauttaa HTML-palvelun sisällön.
 */
function doGet() {
  var title = 'Finance Manager'; // Oletus englanniksi
  
  try {
    var settings = getSettings();
    var lang = settings['Language'];
    
    if (lang === 'fi') title = 'Taloushallinta';
    else if (lang === 'sv') title = 'Ekonomihantering';
    
  } catch (e) {
    console.log("Error fetching title settings: " + e);
  }

  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle(title)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getInitData() {
  var settings = getSettings();
  var translations = getTranslations();
  var categories = getCategories(); 
  
  return {
    settings: settings,
    translations: translations,
    categories: categories
  };
}

// Hakee datan ja laskee Dashboardin tiedot
function getData(year, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(year.toString());

  if (!sheet) {
    return { error: 'Year not found in database. Create it?', year: year };
  }
  
  var dataRange = sheet.getDataRange();
  if (dataRange.getLastRow() < 2) {
    return { 
        year: year, 
        month: month, 
        bills: [], 
        monthlyStats: createEmptyMonthlyStats(),
        annualSummary: createEmptyAnnualSummary() 
    };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var entries = data.slice(1);
  var bills = [];
  var monthlyStats = createEmptyMonthlyStats(); 
  var annualSummary = createEmptyAnnualSummary(); 

  entries.forEach(function(row) {
    var entry = {};
    headers.forEach(function(header, index) {
      entry[header] = row[index];
    });
    
    if (!entry.ID) return;
    
    var amount = parseFloat(entry['Summa']) || 0;
    var rawDate = entry['Eräpäivä'];
    var monthNum = (rawDate instanceof Date) ? rawDate.getMonth() + 1 : parseInt(entry['Kuukausi']);

    if (month > 0 && monthNum !== month) {
      return;
    }

    bills.push({
      id: entry.ID,
      name: entry['Nimi'],
      amount: amount.toFixed(2),
      dueDate: (rawDate instanceof Date) ? Utilities.formatDate(rawDate, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : 'N/A',
      paid: entry['Maksettu'] === true,
      type: entry['Tyyppi'] || 'expense', 
      category: entry['Kategoria'] || 'Ei kategoriaa', 
      recurringId: entry['RecurringID'] || null,
      copyNext: entry['KopioiSeuraavaan'] === true
    });
  });

  var totalIncome = 0;
  var totalExpense = 0;
  var categoryBreakdown = {};
  var expenseMonths = 0;
  var monthsWithExpense = new Set();
    
  entries.forEach(function(row) {
      var entry = {};
      headers.forEach(function(header, index) {
          entry[header] = row[index];
      });
      
      var amount = parseFloat(entry['Summa']) || 0;
      var rawDate = entry['Eräpäivä'];
      var monthNum = (rawDate instanceof Date) ? rawDate.getMonth() + 1 : parseInt(entry['Kuukausi']);
      var type = entry['Tyyppi'] || 'expense';
      var category = entry['Kategoria'] || 'Ei kategoriaa';

      if (monthNum >= 1 && monthNum <= 12) {
          var index = monthNum - 1;
          
          if (type === 'income') {
              monthlyStats[index].income += amount;
              totalIncome += amount;
          } else if (type === 'expense' || type === 'extra') {
              monthlyStats[index].expense += amount;
              totalExpense += amount;
              monthsWithExpense.add(monthNum);

              if (type === 'expense') {
                  monthlyStats[index].billCount++;
                  if (!categoryBreakdown[category]) {
                      categoryBreakdown[category] = 0;
                  }
                  categoryBreakdown[category] += amount;
              } else if (type === 'extra') {
                  var extraCat = 'Budjetti / Muut';
                   if (!categoryBreakdown[extraCat]) {
                      categoryBreakdown[extraCat] = 0;
                  }
                  categoryBreakdown[extraCat] += amount;
              }
          }
      }
  });
  
  expenseMonths = monthsWithExpense.size;
  var averageMonthlyExpense = expenseMonths > 0 ? totalExpense / expenseMonths : 0;
  
  var sortedCategories = Object.keys(categoryBreakdown).map(function(key) {
      return { name: key, total: categoryBreakdown[key] };
  }).sort(function(a, b) {
      return b.total - a.total;
  });

  annualSummary = {
      totalIncome: totalIncome.toFixed(2),
      totalExpense: totalExpense.toFixed(2),
      projectedBalance: (totalIncome - totalExpense).toFixed(2),
      averageMonthlyExpense: averageMonthlyExpense.toFixed(2),
      categoryBreakdown: categoryBreakdown,
      topCategories: sortedCategories.slice(0, 5) 
  };
  
  if (month > 0) {
      monthlyStats = [];
      annualSummary = createEmptyAnnualSummary();
  }

  return {
    year: year,
    month: month,
    bills: bills,
    monthlyStats: monthlyStats,
    annualSummary: annualSummary
  };
}

function createEmptyMonthlyStats() {
    var stats = [];
    for (var i = 1; i <= 12; i++) {
        stats.push({ income: 0, expense: 0, billCount: 0 });
    }
    return stats;
}

function createEmptyAnnualSummary() {
    return {
        totalIncome: '0.00',
        totalExpense: '0.00',
        projectedBalance: '0.00',
        averageMonthlyExpense: '0.00',
        categoryBreakdown: {},
        topCategories: []
    };
}


// --- TIETOJEN HALLINTA ---

function saveTransaction(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = new Date(formData.date).getFullYear();
  var sheet = ss.getSheetByName(year.toString());
  
  if (!sheet) {
    return { success: false, message: 'Taulukkoa vuodelle ' + year + ' ei löydy.' };
  }
  
  var newId = Utilities.getUuid();
  var amount = parseFloat(formData.amount);
  var type = formData.type;
  var category = formData.category || ''; 
  var isRecurring = formData.repeats !== '1';
  var copyToNextYear = formData.copyToNextYear;
  
  var months = [];
  if (isRecurring) {
    if (formData.repeats === 'custom' && formData.customMonths) {
      months = formData.customMonths;
    } else if (formData.repeats !== 'custom') {
      var repeatInterval = parseInt(formData.repeats);
      
      // JOS "12 Kuukautta (Vuosi)" on valittu, muutetaan väliksi 1,
      // jotta se lisää merkinnän joka kuukaudelle.
      if (repeatInterval === 12) {
          repeatInterval = 1;
      }

      var startMonth = new Date(formData.date).getMonth() + 1;
      for (var m = startMonth; m <= 12; m += repeatInterval) {
        months.push(m.toString());
      }
    }
  } else {
    months.push((new Date(formData.date).getMonth() + 1).toString());
  }

  months.forEach(function(month) {
    var entryMonth = parseInt(month);
    var date = new Date(formData.date);
    date.setMonth(entryMonth - 1);
    
    sheet.appendRow([
      newId, 
      formData.name, 
      date, 
      amount, 
      false, 
      isRecurring ? newId : null, 
      entryMonth,
      new Date(), 
      type,
      category, 
      copyToNextYear 
    ]);
  });
  
  return { success: true, message: 'Saved!' };
}

function editTransaction(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = new Date(formData.date).getFullYear();
  var sheet = ss.getSheetByName(year.toString());

  if (!sheet) {
    return { success: false, message: 'Taulukkoa vuodelle ' + year + ' ei löydy.' };
  }

  var data = sheet.getDataRange().getValues();
  var rowToUpdate = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === formData.id) {
      rowToUpdate = i + 1;
      break;
    }
  }

  if (rowToUpdate !== -1) {
    var date = new Date(formData.date);
    var amount = parseFloat(formData.amount);
    var category = formData.category || ''; 
    
    sheet.getRange(rowToUpdate, 2).setValue(formData.name); 
    sheet.getRange(rowToUpdate, 3).setValue(date); 
    sheet.getRange(rowToUpdate, 4).setValue(amount); 
    sheet.getRange(rowToUpdate, 7).setValue(date.getMonth() + 1); 
    sheet.getRange(rowToUpdate, 9).setValue(formData.type); 
    sheet.getRange(rowToUpdate, 10).setValue(category); 
    sheet.getRange(rowToUpdate, 11).setValue(formData.copyToNextYear); 
    
    return { success: true, message: 'Updated!' };
  }
  
  return { success: false, message: 'Entry not found.' };
}

function deleteTransaction(id, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(year.toString());

  if (!sheet) {
    return { success: false, message: 'Sheet not found.' };
  }

  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === id || data[i][5] === id) { 
      rowsToDelete.push(i + 1); 
    }
  }
  
  rowsToDelete.sort(function(a, b){return b - a;});

  rowsToDelete.forEach(function(row) {
    sheet.deleteRow(row);
  });
  
  return { success: true, message: 'Deleted.' };
}

function toggleBillStatus(id, year, isPaid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(year.toString());

  if (!sheet) {
    return { success: false, message: 'Sheet not found.' };
  }

  var data = sheet.getDataRange().getValues();
  var rowToUpdate = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      rowToUpdate = i + 1;
      break;
    }
  }

  if (rowToUpdate !== -1) {
    sheet.getRange(rowToUpdate, 5).setValue(isPaid);
    return { success: true, message: 'Updated.' };
  }
  
  return { success: false, message: 'Entry not found.' };
}

function createNextYear() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var availableYears = getAvailableYears();
  var lastYear = availableYears.length > 0 ? availableYears[availableYears.length - 1] : new Date().getFullYear();
  var nextYear = lastYear + 1;

  var currentSheet = ss.getSheetByName(lastYear.toString());
  var nextSheet = ss.getSheetByName(nextYear.toString());

  if (!currentSheet) {
    return { success: false, message: 'Edellistä vuotta (' + lastYear + ') ei löydy. Luo se ensin.' };
  }
  
  if (nextSheet) {
    return { success: false, message: 'Vuosi ' + nextYear + ' on jo olemassa.' };
  }

  try {
    nextSheet = ss.insertSheet(nextYear.toString());
    var newHeaders = ['ID', 'Nimi', 'Eräpäivä', 'Summa', 'Maksettu', 'RecurringID', 'Kuukausi', 'Aikaleima', 'Tyyppi', 'Kategoria', 'KopioiSeuraavaan']; 
    nextSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    nextSheet.setFrozenRows(1);
    nextSheet.getRange("C:C").setNumberFormat("yyyy-MM-dd");
    nextSheet.getRange("D:D").setNumberFormat("0.00");
    
    var currentData = currentSheet.getDataRange().getValues();
    var currentHeaders = currentData[0];
    
    var idxCopy = currentHeaders.indexOf('KopioiSeuraavaan');
    var idxCategory = currentHeaders.indexOf('Kategoria');
    var idxType = currentHeaders.indexOf('Tyyppi');
    var idxRecurr = currentHeaders.indexOf('RecurringID');
    
    if (idxCopy === -1) idxCopy = 9; 
    
    var newEntries = [];

    for (var i = 1; i < currentData.length; i++) {
      var row = currentData[i];
      var copyNext = row[idxCopy] === true; 

      if (copyNext) {
        var originalDate = row[2];
        var originalMonth = (originalDate instanceof Date) ? originalDate.getMonth() : row[6] - 1;

        var newDate = new Date(nextYear, originalMonth, originalDate.getDate());
        var newId = Utilities.getUuid(); 
        
        var categoryVal = (idxCategory > -1) ? row[idxCategory] : 'Ei kategoriaa';
        var typeVal = (idxType > -1) ? row[idxType] : 'expense';
        var recVal = (idxRecurr > -1) ? row[idxRecurr] : newId;

        newEntries.push([
          newId,                 
          row[1],                
          newDate,               
          row[3],                
          false,                 
          recVal || newId,       
          newDate.getMonth() + 1, 
          new Date(),            
          typeVal,   
          categoryVal,           
          true                   
        ]);
      }
    }
    
    if (newEntries.length > 0) {
      nextSheet.getRange(nextSheet.getLastRow() + 1, 1, newEntries.length, 11).setValues(newEntries);
    }

    return { success: true, message: 'Vuosi ' + nextYear + ' luotu onnistuneesti!' };

  } catch (e) {
    return { success: false, message: 'Virhe vuoden luomisessa: ' + e.toString() };
  }
}

function getAvailableYears() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var years = [];
  
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (/^\d{4}$/.test(name)) {
      years.push(parseInt(name));
    }
  });
  
  years.sort(function(a, b){return a - b;});
  
  return years;
}

// --- KIELI JA ASETUKSET ---

function getSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  var settings = {};
  
  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        data.forEach(function(row) {
          if (row[0]) {
            settings[row[0]] = row[1];
          }
        });
    }
  }
  return settings;
}

function saveSetting(key, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME, 0); 
    sheet.getRange('A1:B1').setValues([['Key', 'Value']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  var data = sheet.getDataRange().getValues();
  var keyFound = false;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      keyFound = true;
      break;
    }
  }
  
  if (!keyFound) {
    sheet.appendRow([key, value]);
  }
}

function getCategories() {
  var settings = getSettings();
  var categoriesString = settings[CATEGORY_SETTING_KEY] || '';
  
  return categoriesString.split(',').map(function(c) { return c.trim(); }).filter(function(c) { return c.length > 0; });
}

function saveCategories(categoriesArray) {
  var categoriesString = categoriesArray.join(', ');
  saveSetting(CATEGORY_SETTING_KEY, categoriesString);
}

function getTranslations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRANSLATIONS_SHEET_NAME);
  var translations = {};
  
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      var headers = data[0]; 
      var keys = headers.slice(1); 
      
      keys.forEach(function(lang, langIndex) {
        translations[lang] = {};
        for (var i = 1; i < data.length; i++) {
          if (data[i][0]) { 
            translations[lang][data[i][0]] = data[i][langIndex + 1];
          }
        }
      });
    }
  }
  return translations;
}

function forceUpdateTranslations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRANSLATIONS_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(TRANSLATIONS_SHEET_NAME);
    sheet.appendRow(['Key', 'fi', 'en', 'sv']);
    sheet.setFrozenRows(1);
  }

  var headers = sheet.getRange(1, 1, 1, 4).getValues()[0];
  if (headers[0] !== 'Key') {
    sheet.getRange(1, 1, 1, 4).setValues([['Key', 'fi', 'en', 'sv']]);
  }

  var existingData = sheet.getDataRange().getValues();
  var existingKeys = [];
  if (existingData.length > 1) {
    existingKeys = existingData.slice(1).map(function(r) { return r[0]; }).filter(String); 
  }
  
  var defaults = [
    ['login_welcome', 'Tervetuloa', 'Welcome', 'Välkommen'],
    ['login_prompt', 'Syötä pääsykoodi', 'Enter Access Code', 'Ange åtkomstkod'],
    ['login_btn', 'Kirjaudu', 'Login', 'Logga in'],
    ['login_placeholder', 'Salasana', 'Password', 'Lösenord'],
    ['login_error', 'Väärä salasana', 'Incorrect Password', 'Fel lösenord'],
    ['app_title', 'Taloushallinta', 'Finance Manager', 'Ekonomihantering'],
    ['btn_new_year', 'Uusi Vuosi', 'New Year', 'Nytt År'],
    ['year_summary', 'Vuosiyhteenveto', 'Annual Summary', 'Årsöversikt'],
    ['financial_overview', 'Talouskatsaus', 'Financial Overview', 'Ekonomisk översikt'],
    ['income_breakdown', 'Tulojen erittely', 'Income Breakdown', 'Inkomstfördelning'],
    ['label_total_income', 'Tulot yhteensä:', 'Total Income:', 'Total inkomst:'],
    ['label_paid', 'Maksetut Laskut:', 'Paid Bills:', 'Betalda räkningar:'],
    ['label_unpaid', 'Maksamattomat Laskut:', 'Unpaid Bills:', 'Obetalda räkningar:'],
    ['label_balance', 'Jäljellä:', 'Balance:', 'Saldo:'],
    ['bills_title', 'Laskut', 'Bills', 'Räkningar'],
    ['drag_tip', 'Vinkki: Raahaa laskuja muuttaaksesi tilaa', 'Tip: Drag bills to change status', 'Tips: Dra räkningar för att ändra status'],
    ['header_paid', 'Maksetut', 'Paid', 'Betalda'],
    ['header_unpaid', 'Maksamatta', 'Unpaid', 'Obetalda'],
    ['modal_add_title', 'Lisää uusi', 'Add New Entry', 'Lägg till ny'],
    ['modal_edit_title', 'Muokkaa', 'Edit Entry', 'Redigera'],
    ['type_expense', 'Lasku', 'Bill', 'Räkning'],
    ['type_income', 'Tulo', 'Income', 'Inkomst'],
    ['label_name', 'Nimi / Kuvaus', 'Name / Description', 'Namn / Beskrivning'],
    ['label_amount', 'Määrä', 'Amount', 'Belopp'],
    ['label_date', 'Eräpäivä', 'Due Date', 'Förfallodatum'],
    ['label_recurrence', 'Toistuvuus', 'Recurrence', 'Återkommande'],
    ['rec_once', 'Kerran', 'Once', 'En gång'],
    ['rec_2m', '2 Kuukautta', '2 Months', '2 Månader'],
    ['rec_3m', '3 Kuukautta', '3 Months', '3 Månader'],
    ['rec_6m', '6 Kuukautta', '6 Months', '6 Månader'],
    ['rec_12m', '12 Kuukautta (Vuosi)', '12 Months (Year)', '12 Månader (År)'],
    ['rec_custom', 'Valitse kuukaudet...', 'Select Specific Months...', 'Välj månader...'],
    ['label_copy_next', 'Kopioi ensi vuodelle', 'Copy to next year', 'Kopiera till nästa år'],
    ['btn_cancel', 'Peruuta', 'Cancel', 'Avbryt'],
    ['btn_save', 'Tallenna', 'Save', 'Spara'],
    ['btn_close', 'Sulje', 'Close', 'Stäng'],
    ['btn_confirm', 'Vahvista', 'Confirm', 'Bekräfta'],
    ['settings_title', 'Asetukset', 'Settings', 'Inställningar'],
    ['settings_lang', 'Kieli', 'Language', 'Språk'],
    ['settings_currency', 'Valuutta', 'Currency', 'Valuta'],
    ['confirm_title', 'Vahvista', 'Confirm', 'Bekräfta'],
    ['table_month', 'Kuukausi', 'Month', 'Månad'],
    ['table_bills', 'Laskut', 'Bills', 'Räkningar'],
    ['table_exp', 'Menot', 'Expenses', 'Utgifter'],
    ['table_inc', 'Tulot', 'Income', 'Inkomster'],
    ['table_surplus', 'Ylijäämä', 'Surplus', 'Överskott'],
    ['msg_whole_year', 'Koko Vuosi', 'Whole Year', 'Hela året'],
    ['err_year_not_found', 'Vuotta ei löytynyt.', 'Year not found.', 'År hittades inte.'],
    ['err_year_missing_q', 'Vuotta ei löytynyt. Luodaanko se?', 'Year not found. Create it?', 'År hittades inte. Skapa det?'],
    ['msg_create_year_q', 'Luodaanko uusi taulukko ensi vuodelle?', 'Create new sheet for next year?', 'Skapa nytt blad för nästa år?'],
    ['label_source', 'Tulon lähde', 'Source Name', 'Källa'],
    ['label_date_received', 'Vastaanottopäivä', 'Date Received', 'Mottagningsdatum'],
    ['alert_error', 'Virhe', 'Error', 'Fel'],
    ['alert_success', 'Onnistui', 'Success', 'Lyckades'],
    ['month_1', 'Tammikuu', 'January', 'Januari'],
    ['month_2', 'Helmikuu', 'February', 'Februari'],
    ['month_3', 'Maaliskuu', 'March', 'Mars'],
    ['month_4', 'Huhtikuu', 'April', 'April'],
    ['month_5', 'Toukokuu', 'May', 'Maj'],
    ['month_6', 'Kesäkuu', 'June', 'Juni'],
    ['month_7', 'Heinäkuu', 'July', 'Juli'],
    ['month_8', 'Elokuu', 'August', 'Augusti'],
    ['month_9', 'Syyskuu', 'September', 'September'],
    ['month_10', 'Lokakuu', 'October', 'Oktober'],
    ['month_11', 'Marraskuu', 'November', 'November'],
    ['month_12', 'Joulukuu', 'December', 'December'],
    ['header_extras', 'Budjetti / Muut', 'Budget / Extras', 'Budget / Övrigt'],
    ['type_extra', 'Muu', 'Extra', 'Övrigt'],
    ['label_extras_total', 'Muut menot:', 'Other Expenses:', 'Övriga utgifter:'],
    ['label_date_general', 'Päivämäärä', 'Date', 'Datum'],
    ['msg_creating_year', 'Luodaan seuraavaa vuotta, odota. Tässä saattaa kestää minuutteja...', 'Creating next year, please wait. This might take minutes...', 'Skapar nästa år, vänta. Detta kan ta minuter...'],
    ['settings_categories', 'Hallinnoi Kategorioita', 'Manage Categories', 'Hantera Kategorier'],
    ['label_category', 'Kategoria', 'Category', 'Kategori'],
    ['no_category', 'Ei kategoriaa', 'No category', 'Ingen kategori'],
    ['category_hint', 'Syötä kategoriat pilkulla eroteltuna (esim. Ruoka, Asuminen, Huvit).', 'Enter categories separated by commas (e.g., Food, Housing, Fun).', 'Ange kategorier åtskilda med kommatecken (t.ex. Mat, Boende, Nöje).'],
    ['label_projected_balance', 'Odotettu Saldo', 'Projected Balance', 'Förväntat Saldo'],
    ['label_total_expense', 'Menot Yhteensä', 'Total Expense', 'Total Utgift'],
    ['label_avg_monthly_exp', 'Keskimääräinen kuukausimeno', 'Avg Monthly Expense', 'Genomsnittlig månatlig utgift'],
    ['label_top_categories', 'Top 5 Menokategoriat', 'Top 5 Expense Categories', 'Topp 5 Utgiftskategorier'],
    ['label_category_breakdown', 'Kategorioiden Erittely', 'Category Breakdown', 'Kategorifördelning'],
    ['dash_no_data', 'Ei laskettavaa dataa vuodelle.', 'No calculable data for the year.', 'Ingen beräkningsbar data för året.'],
    ['btn_save_categories', 'Tallenna Kategoriat', 'Save Categories', 'Spara Kategorier'],
    // --- SALASANAN KÄÄNNÖKSET ---
    ['settings_change_pwd', 'Vaihda Salasana', 'Change Password', 'Byt Lösenord'],
    ['lbl_current_pwd', 'Nykyinen salasana', 'Current Password', 'Nuvarande lösenord'],
    ['lbl_new_pwd', 'Uusi salasana', 'New Password', 'Nytt lösenord'],
    ['btn_update_pwd', 'Päivitä Salasana', 'Update Password', 'Uppdatera Lösenord'],
    ['msg_pwd_changed', 'Salasana vaihdettu!', 'Password changed!', 'Lösenord ändrat!'],
    ['err_wrong_curr_pwd', 'Nykyinen salasana väärin.', 'Incorrect current password.', 'Fel nuvarande lösenord.']
  ];

  var newRows = [];
  defaults.forEach(function(row) {
    if (!row || !row[0]) return;
    
    if (existingKeys.indexOf(row[0]) === -1) {
      newRows.push(row);
    }
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
  }
}

function checkLogin(kayttajanSalasana) {
  var oikeaSalasana = PropertiesService.getScriptProperties().getProperty('APP_PASSWORD');
  
  if (!oikeaSalasana) {
     return false; 
  }
  
  return kayttajanSalasana === oikeaSalasana;
}

// --- FUNKTIO SALASANAN VAIHTAMISEEN ---
function changePassword(oldPwd, newPwd) {
  var props = PropertiesService.getScriptProperties();
  var current = props.getProperty('APP_PASSWORD');
  
  if (current !== oldPwd) {
    return { success: false };
  }
  
  props.setProperty('APP_PASSWORD', newPwd);
  return { success: true };
}

function checkSetupStatus() {
  var props = PropertiesService.getScriptProperties();
  
  var isSetup = props.getProperty('APP_PASSWORD') ? true : false; 

  return {
    isSetup: isSetup,
    currentYear: new Date().getFullYear()
  };
}

function performWizardSetup(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var yearStr = config.year.toString();
  
  try {
    PropertiesService.getScriptProperties().setProperty('APP_PASSWORD', config.password);

    saveSetting('Language', config.language);
    saveSetting('Currency', config.currency);
    
    var defaultCats = [];
    if (config.language === 'en') {
        defaultCats = ['Food', 'Housing', 'Transport', 'Entertainment', 'Groceries', 'Utilities'];
    } else if (config.language === 'sv') {
        defaultCats = ['Mat', 'Boende', 'Transport', 'Nöje', 'Matvaror', 'Verktyg'];
    } else {
        defaultCats = ['Ruoka', 'Asuminen', 'Liikenne', 'Huvit', 'Ostokset', 'Palvelut'];
    }
    saveCategories(defaultCats); 
    
    forceUpdateTranslations();
    
    var sheet = ss.getSheetByName(yearStr);
    
    if (!sheet) {
      sheet = ss.insertSheet(yearStr);
      var headers = ['ID', 'Nimi', 'Eräpäivä', 'Summa', 'Maksettu', 'RecurringID', 'Kuukausi', 'Aikaleima', 'Tyyppi', 'Kategoria', 'KopioiSeuraavaan']; 
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      
      sheet.getRange("C:C").setNumberFormat("yyyy-MM-dd");
      sheet.getRange("D:D").setNumberFormat("0.00");
    }

    return { success: true, message: 'Setup completed!' };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}
