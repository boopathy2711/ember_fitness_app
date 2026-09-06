/**
 * FUEL LEDGER — Backend (Google Apps Script)
 * Bind this script to the Google Sheet that will act as your database
 * (Extensions > Apps Script from within the Sheet).
 *
 * Deploy: Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone (or "Anyone with the link")
 * Copy the deployment URL into API_URL at the top of index.html.
 */

// ---------- SHEET NAMES ----------
var SHEETS = {
  PROFILE: 'Profile',
  WEIGHT: 'WeightLog',
  FOODDB: 'FoodDatabase',
  PANTRY: 'Pantry',
  FOODLOG: 'FoodLog',
  PLANS: 'MealPlans',
  CONFIG: 'Config'
};

// ---------- ENTRY POINTS ----------
function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  var result;
  try {
    var payload = {};
    if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch (parseErr) { payload = e.parameter; }
    } else {
      payload = e.parameter;
    }
    var action = payload.action || (e.parameter && e.parameter.action);
    ensureSheetsExist();

    switch (action) {
      case 'getProfile': result = getProfile(); break;
      case 'saveProfile': result = saveProfile(payload); break;
      case 'getDashboard': result = getDashboard(payload); break;

      case 'logFood': result = logFood(payload); break;
      case 'getFoodLog': result = getFoodLog(payload); break;
      case 'deleteFoodLog': result = deleteFoodLog(payload); break;

      case 'searchFoodDatabase': result = searchFoodDatabase(payload); break;
      case 'addFoodToDatabase': result = addFoodToDatabase(payload); break;

      case 'getPantry': result = getPantry(); break;
      case 'savePantryItem': result = savePantryItem(payload); break;
      case 'deletePantryItem': result = deletePantryItem(payload); break;

      case 'estimateFoodAI': result = estimateFoodAI(payload); break;
      case 'generateMealPlan': result = generateMealPlan(payload); break;
      case 'getSavedMealPlans': result = getSavedMealPlans(); break;
      case 'deleteMealPlan': result = deleteMealPlan(payload); break;

      case 'logWeight': result = logWeight(payload); break;
      case 'getWeightHistory': result = getWeightHistory(); break;

      case 'getConfig': result = getConfig(); break;
      case 'saveConfig': result = saveConfig(payload); break;

      default: result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- SETUP ----------
function ensureSheetsExist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = {};
  defs[SHEETS.PROFILE] = ['Age','Gender','HeightCm','WeightKg','ActivityLevel','Goal','GoalRateKgPerWeek',
    'TargetCalories','TargetProtein','TargetCarbs','TargetFat','OfficeDays','KitchenAccess','CuisinePref','Allergies','MealsPerDay'];
  defs[SHEETS.WEIGHT] = ['Date','WeightKg'];
  defs[SHEETS.FOODDB] = ['ID','Name','Unit','CaloriesPerUnit','ProteinPerUnit','CarbsPerUnit','FatPerUnit','Tags'];
  defs[SHEETS.PANTRY] = ['ID','ItemName','Quantity','Unit','Notes'];
  defs[SHEETS.FOODLOG] = ['ID','Date','Meal','ItemName','Quantity','Calories','Protein','Carbs','Fat','Source'];
  defs[SHEETS.PLANS] = ['ID','DateGenerated','Title','Tag','MealsJSON','TotalCalories','TotalProtein','TotalCarbs','TotalFat'];
  defs[SHEETS.CONFIG] = ['Key','Value'];

  Object.keys(defs).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(defs[name]);
      sh.setFrozenRows(1);
    }
  });

  // Seed a starter FoodDatabase if empty (common Indian + generic staples)
  var fdb = ss.getSheetByName(SHEETS.FOODDB);
  if (fdb.getLastRow() < 2) {
    var starter = [
      ['1','Roti (whole wheat)','piece',80,3,15,1,'veg,indian,staple'],
      ['2','Cooked rice','cup',205,4,45,0.5,'veg,indian,staple'],
      ['3','Dal (cooked)','cup',230,18,40,1,'veg,indian,protein'],
      ['4','Paneer','100g',265,18,4,20,'veg,indian,protein'],
      ['5','Boiled egg','piece',78,6,0.6,5,'nonveg,protein'],
      ['6','Chicken breast (cooked)','100g',165,31,0,3.6,'nonveg,protein'],
      ['7','Banana','piece',105,1.3,27,0.4,'veg,fruit'],
      ['8','Peanut butter','tbsp',95,4,3,8,'veg,fat'],
      ['9','Milk (full fat)','cup',150,8,12,8,'veg,dairy'],
      ['10','Curd/Yogurt','cup',150,8,11,8,'veg,dairy'],
      ['11','Almonds','10 pieces',70,2.5,2.5,6,'veg,fat,snack'],
      ['12','Oats (dry)','cup',307,11,55,5,'veg,carb'],
      ['13','Sweet potato (boiled)','100g',86,1.6,20,0.1,'veg,carb'],
      ['14','Mixed vegetable curry','cup',150,4,18,7,'veg,indian']
    ];
    fdb.getRange(2, 1, starter.length, starter[0].length).setValues(starter);
  }
}

// ---------- PROFILE ----------
function getProfile() {
  var sh = sheet(SHEETS.PROFILE);
  if (sh.getLastRow() < 2) return { success: true, profile: null };
  var row = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var profile = {};
  headers.forEach(function (h, i) { profile[h] = row[i]; });
  return { success: true, profile: profile };
}

function saveProfile(p) {
  var sh = sheet(SHEETS.PROFILE);
  var targets = computeTargets(p);
  var row = [p.age, p.gender, p.heightCm, p.weightKg, p.activityLevel, p.goal, p.goalRate || 0.5,
    targets.calories, targets.protein, targets.carbs, targets.fat,
    p.officeDays || '', p.kitchenAccess || 'none', p.cuisinePref || '', p.allergies || '', p.mealsPerDay || 4];
  if (sh.getLastRow() < 2) {
    sh.appendRow(row);
  } else {
    sh.getRange(2, 1, 1, row.length).setValues([row]);
  }
  // also log current weight into WeightLog for the trend chart
  logWeight({ weightKg: p.weightKg, date: todayStr() });
  return { success: true, targets: targets };
}

function computeTargets(p) {
  var age = Number(p.age), h = Number(p.heightCm), w = Number(p.weightKg);
  var bmr = p.gender === 'female'
    ? (10 * w + 6.25 * h - 5 * age - 161)
    : (10 * w + 6.25 * h - 5 * age + 5);
  var mult = { sedentary: 1.2, light: 1.375, moderate: 1.55, active: 1.725, very_active: 1.9 }[p.activityLevel] || 1.375;
  var tdee = bmr * mult;
  var rateKg = Number(p.goalRate || 0.5); // kg per week
  var dailyDelta = (rateKg * 7700) / 7; // ~7700 kcal per kg of body mass
  var calories;
  if (p.goal === 'gain') calories = tdee + dailyDelta;
  else if (p.goal === 'lose') calories = tdee - dailyDelta;
  else calories = tdee;
  calories = Math.round(calories / 10) * 10;

  var proteinPerKg = p.goal === 'gain' ? 1.8 : (p.goal === 'lose' ? 2.0 : 1.6);
  var protein = Math.round(w * proteinPerKg);
  var fatCals = calories * 0.25;
  var fat = Math.round(fatCals / 9);
  var proteinCals = protein * 4;
  var carbCals = calories - proteinCals - fatCals;
  var carbs = Math.round(Math.max(carbCals, 0) / 4);

  return { calories: calories, protein: protein, carbs: carbs, fat: fat, bmr: Math.round(bmr), tdee: Math.round(tdee) };
}

// ---------- DASHBOARD ----------
function getDashboard(p) {
  var date = (p && p.date) || todayStr();
  var profile = getProfile().profile;
  var logRows = readAll(SHEETS.FOODLOG).filter(function (r) { return r.Date === date; });
  var totals = logRows.reduce(function (acc, r) {
    acc.calories += Number(r.Calories) || 0;
    acc.protein += Number(r.Protein) || 0;
    acc.carbs += Number(r.Carbs) || 0;
    acc.fat += Number(r.Fat) || 0;
    return acc;
  }, { calories: 0, protein: 0, carbs: 0, fat: 0 });

  return { success: true, profile: profile, totals: totals, log: logRows };
}

// ---------- FOOD LOG ----------
function logFood(p) {
  var sh = sheet(SHEETS.FOODLOG);
  var id = Utilities.getUuid();
  sh.appendRow([id, p.date || todayStr(), p.meal || 'snack', p.itemName, p.quantity || 1,
    p.calories || 0, p.protein || 0, p.carbs || 0, p.fat || 0, p.source || 'manual']);
  return { success: true, id: id };
}

function getFoodLog(p) {
  var date = (p && p.date) || todayStr();
  var rows = readAll(SHEETS.FOODLOG).filter(function (r) { return r.Date === date; });
  return { success: true, log: rows };
}

function deleteFoodLog(p) { return deleteRowById(SHEETS.FOODLOG, p.id); }

// ---------- FOOD DATABASE ----------
function searchFoodDatabase(p) {
  var q = (p.query || '').toLowerCase();
  var rows = readAll(SHEETS.FOODDB);
  if (q) rows = rows.filter(function (r) { return (r.Name + ' ' + r.Tags).toLowerCase().indexOf(q) !== -1; });
  return { success: true, items: rows };
}

function addFoodToDatabase(p) {
  var sh = sheet(SHEETS.FOODDB);
  var id = Utilities.getUuid();
  sh.appendRow([id, p.name, p.unit, p.calories, p.protein, p.carbs, p.fat, p.tags || '']);
  return { success: true, id: id };
}

// ---------- PANTRY ----------
function getPantry() { return { success: true, items: readAll(SHEETS.PANTRY) }; }

function savePantryItem(p) {
  var sh = sheet(SHEETS.PANTRY);
  if (p.id) {
    var rowIdx = findRowById(SHEETS.PANTRY, p.id);
    if (rowIdx > -1) {
      sh.getRange(rowIdx, 1, 1, 5).setValues([[p.id, p.itemName, p.quantity, p.unit, p.notes || '']]);
      return { success: true, id: p.id };
    }
  }
  var id = Utilities.getUuid();
  sh.appendRow([id, p.itemName, p.quantity, p.unit, p.notes || '']);
  return { success: true, id: id };
}

function deletePantryItem(p) { return deleteRowById(SHEETS.PANTRY, p.id); }

// ---------- WEIGHT ----------
function logWeight(p) {
  var sh = sheet(SHEETS.WEIGHT);
  sh.appendRow([p.date || todayStr(), p.weightKg]);
  return { success: true };
}

function getWeightHistory() { return { success: true, history: readAll(SHEETS.WEIGHT) }; }

// ---------- CONFIG ----------
function getConfig() {
  var rows = readAll(SHEETS.CONFIG);
  var cfg = {};
  rows.forEach(function (r) { cfg[r.Key] = r.Value; });
  // never echo the raw key back in full for display purposes; frontend just needs to know it's set
  return { success: true, hasGeminiKey: !!cfg.GeminiApiKey, config: cfg };
}

function saveConfig(p) {
  var sh = sheet(SHEETS.CONFIG);
  Object.keys(p).forEach(function (key) {
    if (key === 'action') return;
    var rowIdx = findRowByColValue(SHEETS.CONFIG, 'Key', key);
    if (rowIdx > -1) sh.getRange(rowIdx, 2).setValue(p[key]);
    else sh.appendRow([key, p[key]]);
  });
  return { success: true };
}

// ---------- AI: NATURAL LANGUAGE FOOD ESTIMATION ----------
function estimateFoodAI(p) {
  var cfg = getConfig().config;
  var apiKey = cfg.GeminiApiKey;
  if (!apiKey) return { success: false, error: 'No Gemini API key set in Profile > Settings yet.' };

  var prompt = 'You are a nutrition estimator. Given this food description, return ONLY valid JSON ' +
    '(no markdown fences, no commentary) matching this shape:\n' +
    '{"items":[{"name":string,"quantity":string,"calories":number,"protein":number,"carbs":number,"fat":number}]}\n' +
    'Estimate realistic values for a typical serving. Description: "' + p.text + '"';

  var text = callGemini(prompt, apiKey);
  var parsed = safeParseJson(text);
  if (!parsed) return { success: false, error: 'Could not parse AI response.' };
  return { success: true, items: parsed.items || [] };
}

// ---------- AI: MEAL PLAN GENERATION ----------
function generateMealPlan(p) {
  var cfg = getConfig().config;
  var apiKey = cfg.GeminiApiKey;
  var pantry = readAll(SHEETS.PANTRY);
  var profile = getProfile().profile || {};

  if (!apiKey) {
    return { success: true, source: 'rule-based', plan: ruleBasedMealPlan(p, pantry, profile) };
  }

  var pantryList = pantry.map(function (i) { return i.ItemName + ' (' + i.Quantity + ' ' + i.Unit + ')'; }).join(', ') || 'none listed';
  var officeNote = p.officeDay
    ? 'This is an office working day: prefer low-prep, portable, or no-cook / microwave-only items suitable for eating at a desk.'
    : 'This is a day at home: full cooking is fine.';
  var pantryNote = p.useOnlyPantry
    ? 'Use ONLY the pantry items listed below — do not suggest anything not in this list.'
    : 'Prefer pantry items where possible, but you may suggest other common items too.';

  var prompt = 'You are a nutrition and meal planning assistant. Build a one-day meal plan.\n' +
    'Target for the day: ' + p.targetCalories + ' kcal, ' + p.targetProtein + 'g protein, ' +
    p.targetCarbs + 'g carbs, ' + p.targetFat + 'g fat.\n' +
    'Number of meals: ' + (p.mealsPerDay || 4) + '.\n' +
    'Cuisine preference: ' + (profile.CuisinePref || 'no strong preference') + '.\n' +
    'Allergies/avoid: ' + (profile.Allergies || 'none') + '.\n' +
    officeNote + '\n' + pantryNote + '\n' +
    'Available pantry items: ' + pantryList + '.\n' +
    'Return ONLY valid JSON, no markdown fences, matching:\n' +
    '{"meals":[{"name":string,"items":[string],"calories":number,"protein":number,"carbs":number,"fat":number,"prepMinutes":number}],' +
    '"totalCalories":number,"totalProtein":number,"totalCarbs":number,"totalFat":number,"notes":string}';

  var text = callGemini(prompt, apiKey);
  var parsed = safeParseJson(text);
  if (!parsed) return { success: true, source: 'rule-based', plan: ruleBasedMealPlan(p, pantry, profile) };

  // Save it
  var sh = sheet(SHEETS.PLANS);
  var id = Utilities.getUuid();
  sh.appendRow([id, todayStr(), p.officeDay ? 'Office day plan' : 'Home day plan',
    p.officeDay ? 'office' : 'home', JSON.stringify(parsed), parsed.totalCalories, parsed.totalProtein, parsed.totalCarbs, parsed.totalFat]);

  return { success: true, source: 'ai', plan: parsed, id: id };
}

// Simple greedy fallback when no AI key is configured yet
function ruleBasedMealPlan(p, pantry, profile) {
  var db = readAll(SHEETS.FOODDB);
  var pool = db.slice();
  if (pantry.length && p.useOnlyPantry) {
    var names = pantry.map(function (i) { return i.ItemName.toLowerCase(); });
    pool = db.filter(function (f) { return names.indexOf(f.Name.toLowerCase()) !== -1; });
  }
  if (!pool.length) pool = db;

  var mealsCount = Number(p.mealsPerDay) || 4;
  var perMealTarget = Number(p.targetCalories) / mealsCount;
  var meals = [];
  var totals = { calories: 0, protein: 0, carbs: 0, fat: 0 };

  for (var m = 0; m < mealsCount; m++) {
    var items = [], cals = 0, prot = 0, carb = 0, fat = 0;
    var shuffled = pool.slice().sort(function () { return Math.random() - 0.5; });
    for (var i = 0; i < shuffled.length && cals < perMealTarget * 0.9; i++) {
      var f = shuffled[i];
      items.push(f.Name);
      cals += Number(f.CaloriesPerUnit);
      prot += Number(f.ProteinPerUnit);
      carb += Number(f.CarbsPerUnit);
      fat += Number(f.FatPerUnit);
      if (items.length >= 4) break;
    }
    meals.push({ name: ['Breakfast', 'Lunch', 'Snack', 'Dinner', 'Second snack'][m] || ('Meal ' + (m + 1)), items: items, calories: Math.round(cals), protein: Math.round(prot), carbs: Math.round(carb), fat: Math.round(fat), prepMinutes: p.officeDay ? 5 : 20 });
    totals.calories += cals; totals.protein += prot; totals.carbs += carb; totals.fat += fat;
  }
  return { meals: meals, totalCalories: Math.round(totals.calories), totalProtein: Math.round(totals.protein), totalCarbs: Math.round(totals.carbs), totalFat: Math.round(totals.fat), notes: 'Rule-based suggestion (add a Gemini API key in Profile > Settings for smarter, AI-personalized plans).' };
}

function getSavedMealPlans() {
  var rows = readAll(SHEETS.PLANS);
  rows.forEach(function (r) { try { r.Meals = JSON.parse(r.MealsJSON); } catch (e) { r.Meals = null; } });
  return { success: true, plans: rows.reverse() };
}

function deleteMealPlan(p) { return deleteRowById(SHEETS.PLANS, p.id); }

// ---------- GEMINI ----------
function callGemini(prompt, apiKey) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
    muteHttpExceptions: true
  };
  var resp = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(resp.getContentText());
  if (data.error) throw new Error(data.error.message);
  return data.candidates[0].content.parts[0].text;
}

function safeParseJson(text) {
  try {
    var cleaned = text.replace(/```json/g, '').replace(/```/g, '').trim();
    return JSON.parse(cleaned);
  } catch (e) { return null; }
}

// ---------- HELPERS ----------
function sheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

function readAll(name) {
  var sh = sheet(name);
  if (sh.getLastRow() < 2) return [];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return values.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function findRowById(name, id) { return findRowByColValue(name, 'ID', id); }

function findRowByColValue(name, col, value) {
  var sh = sheet(name);
  if (sh.getLastRow() < 2) return -1;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var colIdx = headers.indexOf(col);
  var values = sh.getRange(2, colIdx + 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) return i + 2;
  }
  return -1;
}

function deleteRowById(name, id) {
  var rowIdx = findRowById(name, id);
  if (rowIdx > -1) sheet(name).deleteRow(rowIdx);
  return { success: true };
}

function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kolkata', 'yyyy-MM-dd');
}
