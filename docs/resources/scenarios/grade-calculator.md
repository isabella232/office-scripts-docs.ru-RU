---
title: 'Пример сценария Office Scripts: Калькулятор производительности'
description: Пример, который определяет процентные и буквенные оценки для класса учащихся.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 0db6f7c116594f7655bfc0adc8f5a79dbbf2a0af
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700402"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Пример сценария Office Scripts: Калькулятор производительности

В этом сценарии лектор Таллинг каждый из оценок на конце каждого учащегося. Вы ввели оценки для своих назначений и тестов при переходе. Теперь можно определить учащихся "фатес".

Вы разрабатываете сценарий, который суммирует оценки для каждой категории точек. Затем каждый учащийся будет назначать буквенную оценку на основе итогового значения. Чтобы обеспечить точность, вы добавляете пару проверок, чтобы определить, слишком низкие или высокие показатели. Если показатель учащегося меньше нуля или больше возможного значения точки, то сценарий помечает ячейку красной заливкой, а не итоговым баллам учащегося. Это будет ясно указывает, какие записи необходимо проверить. Вы также добавите в оценки некоторые базовые параметры, чтобы можно было быстро просмотреть верхнюю и нижнюю часть класса.

## <a name="scripting-skills-covered"></a>Охваченные навыки работы со сценариями

- Форматирование ячеек
- Проверка ошибок
- Регулярные выражения

## <a name="setup-instructions"></a>Инструкции по настройке

1. Скачайте <a href="grade-calculator.xlsx">граде-Калкулатор. xlsx</a> в свой OneDrive.

2. Откройте книгу с помощью Excel для веб-сайта.

3. На вкладке **Автоматизация** откройте **Редактор кода**.

4. В области задач **Редактор кода** нажмите **новый скрипт** и вставьте следующий скрипт в редактор.

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the number of student record rows.
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let studentsRange = sheet.getUsedRange().load("values, rowCount");
      await context.sync();
      console.log("Total students: " + (studentsRange.rowCount - 1));

      // Clean up any formatting from previous runs of the script.
      studentsRange.clear(Excel.ClearApplyTo.formats);
      studentsRange.getColumn(4).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      await context.sync();

      // Parse the headers for the maximum possible scores for each category.
      // The format is `category (score)`.
      let assignmentsMax = studentsRange.values[0][1].match(/\d+/)[0];
      let midTermMax = studentsRange.values[0][2].match(/\d+/)[0];
      let finalsMax = studentsRange.values[0][3].match(/\d+/)[0];
      console.log("Assignments max score:" + assignmentsMax);
      console.log("Mid-term max score: " + midTermMax);
      console.log("Final max score: " + finalsMax);

      // Look at every student row.
      for (let i = 1; i < studentsRange.values.length; i++) {
        let row = studentsRange.values[i];
        let total = row[1] + row[2] + row[3];
        let valid = true;

        // Look for any records that are too low or too high.
        if (row[1] < 0 || row[1] > assignmentsMax) {
          studentsRange.getCell(i, 1).format.fill.color = "Red";
          valid = false;
        }
        if (row[2] < 0 || row[2] > midTermMax) {
          studentsRange.getCell(i, 2).format.fill.color = "Red";
          valid = false;
        }
        if (row[3] < 0 || row[3] > finalsMax) {
          studentsRange.getCell(i, 3).format.fill.color = "Red";
          valid = false;
        }

        // If the scores are valid, total that student's points and assign them a letter grade.
        if (valid) {
          let grade: string;
          switch (true) {
            case total < 60:
              grade = "E";
              break;
            case total < 70:
              grade = "D";
              break;
            case total < 80:
              grade = "C";
              break;
            case total < 90:
              grade = "B";
              break;
            default:
              grade = "A";
              break;
          }

          studentsRange.getCell(i, 4).values = [[total]];
          studentsRange.getCell(i, 5).values = [[grade]];

          // Highlight excellent students and those in need of attention.
          if (grade === "A") {
            studentsRange.getCell(i, 5).format.fill.color = "Green";
          } else if (grade === "E" || grade === "D") {
            studentsRange.getCell(i, 5).format.fill.color = "Orange";
          }
        }
      }

      studentsRange.getColumn(5).format.horizontalAlignment = "Center";
    }
    ```

5. Переименуйте сценарий на **Оценка калькулятора** и сохраните его.

## <a name="running-the-script"></a>Выполнение скрипта

Запустите сценарий **калькулятора** на листе. Сценарий выполнит итоговые оценки и присвоит каждому студенте буквенную оценку. Если для какого-либо из конкретных оценок задано больше баллов, чем стоит на назначении или тестировании, то несвязанное с ним помечается красным, а итоговое значение не вычисляется.

### <a name="before-running-the-script"></a>Перед выполнением скрипта

![Лист, показывающий строки оценок для учащихся.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a>После выполнения скрипта

![Лист с данными оценки учащегося с недопустимыми ячейками в красном итоге для допустимых строк учащихся.](../../images/scenario-grade-calculator-after.png)
