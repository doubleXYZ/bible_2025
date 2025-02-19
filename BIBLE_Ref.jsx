// Функция для поиска и замены тэгированных сносок с сохранением форматирования
function convertFootnotesFromTags() {
    // Получаем активный документ
    var doc = app.activeDocument;
    
    // Получаем текст документа
    var myStory = doc.stories.item(0);
    
    // Регулярное выражение для поиска тэгированных сносок
    var regex = /<FootnoteStart:><ref::>(.*?)<::ref>(.*?)<FootnoteEnd:>/g;
    
    // Находим все тэгированные сноски и преобразуем их
    var match;
    while ((match = regex.exec(myStory.contents)) !== null) {
        var referenceText = match[1];
        var footnoteText = match[2];
        
        // Создаем критерии поиска для ссылки на сноску
        var findObj = app.findGrepPreferences;
        var saveFindPreferences = findObj.properties;
        findObj.findWhat = referenceText;
        
        // Создаем критерии замены для удаления тэги
        var changeObj = app.changeGrepPreferences;
        var saveChangePreferences = changeObj.properties;
        changeObj.changeTo = "";
        
        // Находим диапазон текста ссылки на сноску
        var foundItems = myStory.findGrep();
        if (foundItems.length > 0) {
            var refTextRange = foundItems[0].parentTextFrames[0].texts[0].characters;
            
            // Сохраняем форматирование ссылки на сноску
            var refFormatting = refTextRange.getElements()[0].properties;
            
            // Вставляем сноску
            var footnote = doc.footnotes.add(myStory.insertionPoints.item(-1), footnoteText);
            
            // Назначаем стиль символа RefStyle тексту ссылки на сноску
            refTextRange.appliedCharacterStyle = doc.characterStyles.item("RefStyle");
            
            // Применим сохраненное форматирование к сноске
            footnote.paragraphs.item(0).characters.item(0).properties = refFormatting;
        }
        
        // Восстанавливаем предпочтения поиска и замены
        // app.findGrepPreferences = saveFindPreferences;
        // app.changeGrepPreferences = saveChangePreferences;
        
        // Удаляем тэги из текста
        myStory.changeGrep();
    }
}

// Запуск функции
convertFootnotesFromTags();


// var regex = /<FootnoteStart:><ref::>(.*?)<::ref>(.*?)<FootnoteEnd:>/g;