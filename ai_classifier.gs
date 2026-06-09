var API_KEY = "sk-proj-YOUR_API_KEY_HERE"; // Вставьте свой ключ OpenAI

var TARGET_COLUMN = "U"; // Например: "Z" или "CL"

// Обратите внимание: текст обрамлен косыми кавычками ` (находятся на клавише с буквой Ё)
var PROMPT_TEXT = `Ты — классификатор коммерческой недвижимости для розничной торговли. Проанализируй текст и верни «НЕТ», только если это предложение о сдаче целого торгового или офисно-торгового помещения на 1 этаже с прямым доступом к улице. Если есть следующие признаки, то верни «ДА».

– только офисного назначения   
– подвальные, цокольные или антресольные помещения полностью или частично
– имеется  единственный вход только со двора или внутри бизнес/торгового центра  
– сдается часть помещения (половина или отдельные кабинеты, коворкинг, кабинеты врачей, beauty-комнаты) 
– почасовая аренда (зал за час), open space
– помещение не достроено или не введено в эксплуатацию (игнорируй упоминания ремонта)`;

// ==========================================
// 🚀 ОСНОВНОЙ КОД (Ниже ничего менять не нужно)
// ==========================================

function runAiAnalysis() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    var cellText = values[i][0];
    
    // Пропускаем пустые ячейки
    if (!cellText || cellText.toString().trim() === "") continue;

    var payload = {
      "model": "gpt-4o-mini",
      "messages": [{"role": "user", "content": PROMPT_TEXT + "\n\nТекст: " + cellText}],
      "temperature": 0
    };

    var options = {
      "method": "post",
      "headers": {
        "Authorization": "Bearer " + API_KEY,
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    try {
      var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
      var json = JSON.parse(response.getContentText());

      if (json.choices && json.choices.length > 0) {
        var answer = json.choices[0].message.content.trim();
        sheet.getRange(TARGET_COLUMN + (startRow + i)).setValue(answer);
      } else {
        sheet.getRange(TARGET_COLUMN + (startRow + i)).setValue("Ошибка API");
      }
    } catch (e) {
      sheet.getRange(TARGET_COLUMN + (startRow + i)).setValue("Сбой сети: " + e.message);
    }
  }
}

// ==========================================
// 📋 МЕНЮ В ТАБЛИЦЕ
// ==========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 ИИ Анализ')
      .addItem('▶️ Запустить классификацию', 'runAiAnalysis')
      .addToUi();
}
