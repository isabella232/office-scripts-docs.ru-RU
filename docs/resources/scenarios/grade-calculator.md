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
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="21c13-103">Пример сценария Office Scripts: Калькулятор производительности</span><span class="sxs-lookup"><span data-stu-id="21c13-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="21c13-104">В этом сценарии лектор Таллинг каждый из оценок на конце каждого учащегося.</span><span class="sxs-lookup"><span data-stu-id="21c13-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="21c13-105">Вы ввели оценки для своих назначений и тестов при переходе.</span><span class="sxs-lookup"><span data-stu-id="21c13-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="21c13-106">Теперь можно определить учащихся "фатес".</span><span class="sxs-lookup"><span data-stu-id="21c13-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="21c13-107">Вы разрабатываете сценарий, который суммирует оценки для каждой категории точек.</span><span class="sxs-lookup"><span data-stu-id="21c13-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="21c13-108">Затем каждый учащийся будет назначать буквенную оценку на основе итогового значения.</span><span class="sxs-lookup"><span data-stu-id="21c13-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="21c13-109">Чтобы обеспечить точность, вы добавляете пару проверок, чтобы определить, слишком низкие или высокие показатели.</span><span class="sxs-lookup"><span data-stu-id="21c13-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="21c13-110">Если показатель учащегося меньше нуля или больше возможного значения точки, то сценарий помечает ячейку красной заливкой, а не итоговым баллам учащегося.</span><span class="sxs-lookup"><span data-stu-id="21c13-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="21c13-111">Это будет ясно указывает, какие записи необходимо проверить.</span><span class="sxs-lookup"><span data-stu-id="21c13-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="21c13-112">Вы также добавите в оценки некоторые базовые параметры, чтобы можно было быстро просмотреть верхнюю и нижнюю часть класса.</span><span class="sxs-lookup"><span data-stu-id="21c13-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="21c13-113">Охваченные навыки работы со сценариями</span><span class="sxs-lookup"><span data-stu-id="21c13-113">Scripting skills covered</span></span>

- <span data-ttu-id="21c13-114">Форматирование ячеек</span><span class="sxs-lookup"><span data-stu-id="21c13-114">Cell formatting</span></span>
- <span data-ttu-id="21c13-115">Проверка ошибок</span><span class="sxs-lookup"><span data-stu-id="21c13-115">Error checking</span></span>
- <span data-ttu-id="21c13-116">Регулярные выражения</span><span class="sxs-lookup"><span data-stu-id="21c13-116">Regular expressions</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="21c13-117">Инструкции по настройке</span><span class="sxs-lookup"><span data-stu-id="21c13-117">Setup instructions</span></span>

1. <span data-ttu-id="21c13-118">Скачайте <a href="grade-calculator.xlsx">граде-Калкулатор. xlsx</a> в свой OneDrive.</span><span class="sxs-lookup"><span data-stu-id="21c13-118">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="21c13-119">Откройте книгу с помощью Excel для веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="21c13-119">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="21c13-120">На вкладке **Автоматизация** откройте **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="21c13-120">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="21c13-121">В области задач **Редактор кода** нажмите **новый скрипт** и вставьте следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="21c13-121">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="21c13-122">Переименуйте сценарий на **Оценка калькулятора** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="21c13-122">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="21c13-123">Выполнение скрипта</span><span class="sxs-lookup"><span data-stu-id="21c13-123">Running the script</span></span>

<span data-ttu-id="21c13-124">Запустите сценарий **калькулятора** на листе.</span><span class="sxs-lookup"><span data-stu-id="21c13-124">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="21c13-125">Сценарий выполнит итоговые оценки и присвоит каждому студенте буквенную оценку.</span><span class="sxs-lookup"><span data-stu-id="21c13-125">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="21c13-126">Если для какого-либо из конкретных оценок задано больше баллов, чем стоит на назначении или тестировании, то несвязанное с ним помечается красным, а итоговое значение не вычисляется.</span><span class="sxs-lookup"><span data-stu-id="21c13-126">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="21c13-127">Перед выполнением скрипта</span><span class="sxs-lookup"><span data-stu-id="21c13-127">Before running the script</span></span>

![Лист, показывающий строки оценок для учащихся.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="21c13-129">После выполнения скрипта</span><span class="sxs-lookup"><span data-stu-id="21c13-129">After running the script</span></span>

![Лист с данными оценки учащегося с недопустимыми ячейками в красном итоге для допустимых строк учащихся.](../../images/scenario-grade-calculator-after.png)
