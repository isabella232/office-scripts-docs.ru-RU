---
title: 'Пример сценария Office Scripts: Калькулятор производительности'
description: Пример, который определяет процентные и буквенные оценки для класса учащихся.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 6f8e3db756c72cf1d0e2f774ccd819c041f0c42d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878642"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="01c3d-103">Пример сценария Office Scripts: Калькулятор производительности</span><span class="sxs-lookup"><span data-stu-id="01c3d-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="01c3d-104">В этом сценарии лектор Таллинг каждый из оценок на конце каждого учащегося.</span><span class="sxs-lookup"><span data-stu-id="01c3d-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="01c3d-105">Вы ввели оценки для своих назначений и тестов при переходе.</span><span class="sxs-lookup"><span data-stu-id="01c3d-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="01c3d-106">Теперь можно определить учащихся "фатес".</span><span class="sxs-lookup"><span data-stu-id="01c3d-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="01c3d-107">Вы разрабатываете сценарий, который суммирует оценки для каждой категории точек.</span><span class="sxs-lookup"><span data-stu-id="01c3d-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="01c3d-108">Затем каждый учащийся будет назначать буквенную оценку на основе итогового значения.</span><span class="sxs-lookup"><span data-stu-id="01c3d-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="01c3d-109">Чтобы обеспечить точность, вы добавляете пару проверок, чтобы определить, слишком низкие или высокие показатели.</span><span class="sxs-lookup"><span data-stu-id="01c3d-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="01c3d-110">Если показатель учащегося меньше нуля или больше возможного значения точки, то сценарий помечает ячейку красной заливкой, а не итоговым баллам учащегося.</span><span class="sxs-lookup"><span data-stu-id="01c3d-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="01c3d-111">Это будет ясно указывает, какие записи необходимо проверить.</span><span class="sxs-lookup"><span data-stu-id="01c3d-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="01c3d-112">Вы также добавите в оценки некоторые базовые параметры, чтобы можно было быстро просмотреть верхнюю и нижнюю часть класса.</span><span class="sxs-lookup"><span data-stu-id="01c3d-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="01c3d-113">Охваченные навыки работы со сценариями</span><span class="sxs-lookup"><span data-stu-id="01c3d-113">Scripting skills covered</span></span>

- <span data-ttu-id="01c3d-114">Форматирование ячеек</span><span class="sxs-lookup"><span data-stu-id="01c3d-114">Cell formatting</span></span>
- <span data-ttu-id="01c3d-115">Проверка ошибок</span><span class="sxs-lookup"><span data-stu-id="01c3d-115">Error checking</span></span>
- <span data-ttu-id="01c3d-116">Регулярные выражения</span><span class="sxs-lookup"><span data-stu-id="01c3d-116">Regular expressions</span></span>
- <span data-ttu-id="01c3d-117">Условное форматирование</span><span class="sxs-lookup"><span data-stu-id="01c3d-117">Conditional formatting</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="01c3d-118">Инструкции по настройке</span><span class="sxs-lookup"><span data-stu-id="01c3d-118">Setup instructions</span></span>

1. <span data-ttu-id="01c3d-119">Скачайте <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="01c3d-119">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="01c3d-120">Откройте книгу с помощью Excel для веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="01c3d-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="01c3d-121">На вкладке **Автоматизация** откройте **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="01c3d-121">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="01c3d-122">В области задач **Редактор кода** нажмите **новый скрипт** и вставьте следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="01c3d-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the worksheet and validate the data.
      let studentsRange = workbook.getActiveWorksheet().getUsedRange();
      if (studentsRange.getColumnCount() !== 6) {
        throw new Error(`The required columns are not present. Expected column headers: "Student ID | Assignment score | Mid-term | Final | Total | Grade"`);
      }

      let studentData = studentsRange.getValues();

      // Clear the total and grade columns.
      studentsRange.getColumn(4).getCell(1, 0).getAbsoluteResizedRange(studentData.length - 1, 2).clear();

      // Clear all conditional formatting.
      workbook.getActiveWorksheet().getUsedRange().clearAllConditionalFormats();

      // Use regular expressions to read the max score from the assignment, mid-term, and final scores columns.
      let maxScores: string[] = [];
      const assignmentMaxMatches = studentData[0][1].match(/\d+/);
      const midtermMaxMatches = studentData[0][2].match(/\d+/);
      const finalMaxMatches = studentData[0][3].match(/\d+/);

      // Check the matches happened before proceeding.
      if (!(assignmentMaxMatches && midtermMaxMatches && finalMaxMatches)) {
        throw new Error(`The scores are not present in the column headers. Expected format: "Assignments (n)|Mid-term (n)|Final (n)"`);
      }

      // Use the first (and only) match from the regular expressions as the max scores.
      maxScores = [assignmentMaxMatches[0], midtermMaxMatches[0], finalMaxMatches[0]];

      // Set conditional formatting for each of the assignment, mid-term, and final scores columns.
      maxScores.forEach((score, i) => {
        let range = studentsRange.getColumn(i + 1).getCell(0, 0).getRowsBelow(studentData.length - 1);
        setCellValueConditionalFormatting(
          score,
          range,
          "#9C0006",
          "#FFC7CE",
          ExcelScript.ConditionalCellValueOperator.greaterThan
        )
      });

      // Store the current range information to avoid calling the workbook in the loop.
      let studentsRangeFormulas = studentsRange.getColumn(4).getFormulasR1C1();
      let studentsRangeValues = studentsRange.getColumn(5).getValues();

      /* Iterate over each of the student rows and compute the total score and letter grade.
      * Note that iterator starts at index 1 to skip first (header) row.
      */
      for (let i = 1; i < studentData.length; i++) {
        // If any of the scores are invalid, skip processing it.
        if (studentData[i][1] > maxScores[0] ||
          studentData[i][2] > maxScores[1] ||
          studentData[i][3] > maxScores[2]) {
          continue;
        }
        const total = studentData[i][1] + studentData[i][2] + studentData[i][3];
        let grade: string;
        switch (true) {
          case total < 60:
            grade = "F";
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

        // Set total score formula.
        studentsRangeFormulas[i][0] = '=RC[-2]+RC[-1]';
        // Set grade cell.
        studentsRangeValues[i][0] = grade;
      }

      // Set the formulas and values outside the loop.
      studentsRange.getColumn(4).setFormulasR1C1(studentsRangeFormulas);
      studentsRange.getColumn(5).setValues(studentsRangeValues);

      // Put a conditional formatting on the grade column.
      let totalRange = studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentData.length - 1);
      setCellValueConditionalFormatting(
        "A",
        totalRange,
        "#001600",
        "#C6EFCE",
        ExcelScript.ConditionalCellValueOperator.equalTo
      );
      ["D", "F"].forEach((grade) => {
        setCellValueConditionalFormatting(
          grade,
          totalRange,
          "#9C0006",
          "#FFC7CE",
          ExcelScript.ConditionalCellValueOperator.equalTo
        );
      })
      // Center the grade column.
      studentsRange.getColumn(5).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    }

    /**
     * Helper function to apply conditional formatting.
     * @param value Cell value to use in conditional formatting formula1.
     * @param range Target range.
     * @param fontColor Font color to use.
     * @param fillColor Fill color to use.
     * @param operator Operator to use in conditional formatting.
     */
    function setCellValueConditionalFormatting(
      value: string,
      range: ExcelScript.Range,
      fontColor: string,
      fillColor: string,
      operator: ExcelScript.ConditionalCellValueOperator) {
      // Determine the formula1 based on the type of value parameter.
      let formula1: string;
      if (isNaN(Number(value))) {
        // For cell value equalTo rule, use this format: formula1: "=\"A\"",
        formula1 = `=\"${value}\"`;
      } else {
        // For number input (greater-than or less-than rules), just append '='.
        formula1 = `=${value}`;
      }

      // Apply conditional formatting.
      let conditionalFormatting : ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({formula1, operator});
    }
    ```

5. <span data-ttu-id="01c3d-123">Переименуйте сценарий на **Оценка калькулятора** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="01c3d-123">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="01c3d-124">Выполнение скрипта</span><span class="sxs-lookup"><span data-stu-id="01c3d-124">Running the script</span></span>

<span data-ttu-id="01c3d-125">Запустите сценарий **калькулятора** на листе.</span><span class="sxs-lookup"><span data-stu-id="01c3d-125">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="01c3d-126">Сценарий выполнит итоговые оценки и присвоит каждому студенте буквенную оценку.</span><span class="sxs-lookup"><span data-stu-id="01c3d-126">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="01c3d-127">Если для какого-либо из конкретных оценок задано больше баллов, чем стоит на назначении или тестировании, то несвязанное с ним помечается красным, а итоговое значение не вычисляется.</span><span class="sxs-lookup"><span data-stu-id="01c3d-127">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="01c3d-128">Перед выполнением скрипта</span><span class="sxs-lookup"><span data-stu-id="01c3d-128">Before running the script</span></span>

![Лист, показывающий строки оценок для учащихся.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="01c3d-130">После выполнения скрипта</span><span class="sxs-lookup"><span data-stu-id="01c3d-130">After running the script</span></span>

![Лист с данными оценки учащегося с недопустимыми ячейками в красном итоге для допустимых строк учащихся.](../../images/scenario-grade-calculator-after.png)
