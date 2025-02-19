// Регулярное выражение для поиска текста
var regex = /<reff::>(.*?)<::ref>/g;

// Получаем текст документа
var myText = doc.stories.item(0).contents;

// Находим все тэгированные сноски и преобразуем их
var match;
while ((match = regex.exec(myText)) !== null) {
    var referenceText = match[1];
    
    // Ищем текст ссылки на сноску
    var refTextRange = doc.stories.item(0).findText(referenceText)[0];
    
    // Сохраняем форматирование ссылки на сноску
    var refFormatting = refTextRange.getElements()[0].properties;
    
    // Вставляем текст ссылки на сноску рядом со сноской в тексте
    var insertionPoint = footnote.insertionPoints.item(-1);
    insertionPoint.contents = referenceText;
    
    // Применим сохраненное форматирование к тексту ссылки на сноску
    insertionPoint.characters.item(0).properties = refFormatting;
    
    // Удаляем тэги из текста сноски
    doc.stories.item(0).changeText(match[0], "");
}
