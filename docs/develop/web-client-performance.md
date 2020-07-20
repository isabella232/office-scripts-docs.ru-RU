---
title: Повышение производительности сценариев Office
description: Создавайте более быстрые сценарии, чтобы понять связь между книгой Excel и сценарием.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: Auto
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878900"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Повышение производительности сценариев Office

Сценарии Office предназначены для автоматизации часто выполняемого ряда задач для экономии времени. Медленный сценарий может быть очень похожим, так как он не ускоряет рабочий процесс. В большинстве случаев ваш сценарий прекрасно подходит и выполняется должным образом. Однако существует несколько проблем, которые могут повлиять на производительность.

Наиболее распространенная причина медленного сценария заключается в чрезмерном обмене данными с книгой. Ваш сценарий выполняется на локальном компьютере, а книга существует в облаке. В определенное время сценарий синхронизирует свои локальные данные с книгой. Это означает, что любые операции записи (например, `workbook.addWorksheet()` ) применяются к книге только при выполнении этой фоновой синхронизации. Аналогично, любые операции чтения (например, `myRange.getValues()` ) получают данные только из книги для сценария в указанное время. В обоих случаях скрипт получает сведения, прежде чем они действуют с данными. Например, приведенный ниже код будет точно регистрировать количество строк в используемом диапазоне.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

API сценариев Office Обеспечьте точность и актуальность данных в книге или сценарии при необходимости. Для правильной работы сценария не нужно беспокоиться об этих синхронизациях. Тем не менее, сведения о связи между этими сценариями и облаками могут помочь вам избежать ненужных сетевых вызовов.

## <a name="performance-optimizations"></a>Оптимизация производительности

Вы можете применять простые методы для сокращения взаимодействия с облаком. Следующие шаблоны помогают ускорить работу ваших сценариев.

- Считывание данных книги один раз, а не повторно в цикле.
- Удалите лишние `console.log` операторы.
- Не используйте блоки try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Считывание данных книги вне цикла

Любой метод, который получает данные из книги, может активировать сетевой вызов. Вместо многократного совершения такого же вызова следует по возможности сохранять данные локально. Это особенно справедливо при работе с циклами.

Рассмотрим сценарий, чтобы получить количество отрицательных чисел в используемом диапазоне листа. Скрипту необходимо выполнить итерацию по каждой ячейке в используемом диапазоне. Для этого требуется диапазон, количество строк и число столбцов. Перед началом цикла необходимо сохранить их как локальные переменные. В противном случае каждая итерация цикла будет принудительно возвращаться к книге.

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> В качестве эксперимента попробуйте заменить `usedRangeValues` в цикле `usedRange.getValues()` . Вы можете заметить, что сценарий выполняется значительно дольше при работе с большими диапазонами.

### <a name="remove-unnecessary-consolelog-statements"></a>Удаление ненужных `console.log` операторов

Ведение журнала консоли — это крайне важное средство для [отладки сценариев](../testing/troubleshooting.md). Тем не менее, он принудительно выполняет синхронизацию с книгой, чтобы убедиться в том, что зарегистрированные данные актуальны. Перед предоставлением общего доступа к сценарию рекомендуется удалить лишние операторы ведения журнала (например, используемые для тестирования). Это, как правило, не вызывает значительных проблем с производительностью, если `console.log()` оператор не находится в цикле.

### <a name="avoid-using-trycatch-blocks"></a>Не используйте блоки try/catch

Мы не рекомендуем использовать [ `try` / `catch` блоки](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) как часть ожидаемого потокового управления сценария. Большинство ошибок можно избежать, проверив объекты, возвращенные из книги. Например, следующий сценарий проверяет, существует ли таблица, возвращаемая книгой, прежде чем пытаться добавить строку.

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a>Справка по случаю обращения

Так как платформа сценариев Office расширяется для работы со средствами [автоматизации](https://flow.microsoft.com/), [адаптивными картами](https://docs.microsoft.com/adaptive-cards)и другими компонентами, подробные сведения о связи между книгой становятся более сложными. Если вам нужна помощь быстрее при выполнении скрипта, обратитесь к разделу [переполнение стека](https://stackoverflow.com/questions/tagged/office-scripts). Не забудьте пометить свой вопрос знакомым "Office — Scripts", чтобы специалисты могли найти их и помочь.

## <a name="see-also"></a>См. также

- [Основы сценариев для сценариев Office в Excel в Интернете](scripting-fundamentals.md)
- [Веб-документы МДН: циклы и итерация](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
