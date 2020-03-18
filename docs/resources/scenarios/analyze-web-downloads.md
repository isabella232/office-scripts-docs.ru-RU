---
title: 'Сценарий примера сценариев Office: анализ загружаемых веб-файлов'
description: Пример, который принимает необработанные данные из Интернета в книгу Excel и определяет исходное расположение, прежде чем упорядочивать эту информацию в таблице.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 9ee12c8d4ca7c191168e3734d7cd9eadc333c165
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700411"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="92423-103">Сценарий примера сценариев Office: анализ загружаемых веб-файлов</span><span class="sxs-lookup"><span data-stu-id="92423-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="92423-104">В этом сценарии вы являетесь задачей анализа загрузки отчетов с веб-сайта компании.</span><span class="sxs-lookup"><span data-stu-id="92423-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="92423-105">Цель этого анализа — определить, поступает ли веб-трафик из Соединенных Штатов Америки или других стран мира.</span><span class="sxs-lookup"><span data-stu-id="92423-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="92423-106">Ваши коллеги отправляют необработанные данные в вашу книгу.</span><span class="sxs-lookup"><span data-stu-id="92423-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="92423-107">В наборе данных каждой недели есть собственный лист.</span><span class="sxs-lookup"><span data-stu-id="92423-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="92423-108">Кроме того, существует **сводный** лист с таблицей и диаграммой, в которой показаны тенденции за неделю.</span><span class="sxs-lookup"><span data-stu-id="92423-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="92423-109">Вы разрабатываете сценарий, который анализирует еженедельные загрузки данных на активном листе.</span><span class="sxs-lookup"><span data-stu-id="92423-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="92423-110">Он будет анализировать IP-адрес, связанный с каждым загружаемым пакетом, и определять, был ли он передан из США.</span><span class="sxs-lookup"><span data-stu-id="92423-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="92423-111">Ответ будет вставлен на лист в виде логического значения ("TRUE" или "FALSE"), а условное форматирование будет применено к этим ячейкам.</span><span class="sxs-lookup"><span data-stu-id="92423-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="92423-112">Результаты размещения IP-адресов будут суммироваться на листе и скопированы в сводную таблицу.</span><span class="sxs-lookup"><span data-stu-id="92423-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="92423-113">Охваченные навыки работы со сценариями</span><span class="sxs-lookup"><span data-stu-id="92423-113">Scripting skills covered</span></span>

- <span data-ttu-id="92423-114">Синтаксический анализ текста</span><span class="sxs-lookup"><span data-stu-id="92423-114">Text parsing</span></span>
- <span data-ttu-id="92423-115">Подфункции в скриптах</span><span class="sxs-lookup"><span data-stu-id="92423-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="92423-116">Условное форматирование</span><span class="sxs-lookup"><span data-stu-id="92423-116">Conditional formatting</span></span>
- <span data-ttu-id="92423-117">таблицы;</span><span class="sxs-lookup"><span data-stu-id="92423-117">Tables</span></span>

## <a name="demo-video"></a><span data-ttu-id="92423-118">Демонстрационное видео</span><span class="sxs-lookup"><span data-stu-id="92423-118">Demo video</span></span>

<span data-ttu-id="92423-119">В этом примере показана демонстрация при вызове сообщества разработчиков надстроек Office в течение февраля 2020.</span><span class="sxs-lookup"><span data-stu-id="92423-119">This sample was demoed as part of the Office Add-ins developer community call for February 2020.</span></span>

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

## <a name="setup-instructions"></a><span data-ttu-id="92423-120">Инструкции по настройке</span><span class="sxs-lookup"><span data-stu-id="92423-120">Setup instructions</span></span>

1. <span data-ttu-id="92423-121">Скачайте <a href="analyze-web-downloads.xlsx">анализе-веб-довнлоадс. xlsx</a> в свой OneDrive.</span><span class="sxs-lookup"><span data-stu-id="92423-121">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="92423-122">Откройте книгу с помощью Excel для веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="92423-122">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="92423-123">На вкладке **Автоматизация** откройте **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="92423-123">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="92423-124">В области задач **Редактор кода** нажмите **новый скрипт** и вставьте следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="92423-124">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
      async function main(context: Excel.RequestContext) {
        let currentWorksheet = context.workbook.worksheets
          .getActiveWorksheet();
        // Get the values of the active range of the active worksheet.
        let logRange = currentWorksheet.getUsedRange().load("values");

        // Get the Summary worksheet and table.
        let summaryWorksheet = context.workbook.worksheets.getItem("Summary");
        let summaryTable = context.workbook.tables.getItem("Table1");

        // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
        let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1)
          .load("address");

        // Get the values of all the US IP addresses.
        let ipRange = context.workbook.worksheets
          .getItem("USIPAddresses")
          .getUsedRange()
          .load("values");
        await context.sync();

        // Remove the first row.
        let topRow = logRange.values.shift();

        // Create a new array to contain the boolean representing if this is a US IP address.
        let newCol = [[]];

        // Go through each row in worksheet and add Boolean.
        for (let i = 0; i < logRange.values.length; i++) {
          let curRowIP = logRange.values[i][1];
          if (findIP(ipRange.values, ipAddressToInteger(curRowIP)) > 0) {
            newCol.push([true]);
          } else {
            newCol.push([false]);
          }
        }

        // Remove the empty column header and add proper heading.
        newCol.shift();
        newCol.unshift(["Is US IP"]);

        // Write the result to the spreadsheet.
        isUSColumn.values = newCol;
        addSummaryData();
        applyConditionalFormatting();
        currentWorksheet.getUsedRange().format.autofitColumns();

        // Get the calculated summary data.
        let summaryRange = currentWorksheet.getRange("J2:M2").load("values");
        await context.sync();

        // Add the corresponding row to the summary table.
        summaryTable.rows.add(null, summaryRange.values);

        // Function to apply conditional formatting to the new column.
        function applyConditionalFormatting() {
          // Add conditional formatting to the new column.
          let conditionalFormatTrue = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          let conditionalFormatFalse = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          // Set TRUE to light blue and FALSE to light orange.
          conditionalFormatTrue.cellValue.format.fill.color = "#8FA8DB";
          conditionalFormatTrue.cellValue.rule = {
            formula1: "=TRUE",
            operator: "EqualTo"
          };
          conditionalFormatFalse.cellValue.format.fill.color = "#F8CCAD";
          conditionalFormatFalse.cellValue.rule = {
            formula1: "=FALSE",
            operator: "EqualTo"
          };
        }

        // Adds the summary data to the current sheet and to the summary table.
        function addSummaryData() {
          // Add a summary row and table.
          let summaryHeader = [["Year", "Week", "US", "Other"]];
          let countTrueFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=TRUE")/' + (newCol.length - 1);
          let countFalseFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=FALSE")/' + (newCol.length - 1);

          let summaryContent = [
            [
              '=TEXT(A2,"YYYY")',
              '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
              countTrueFormula,
              countFalseFormula
            ]
          ];
          let summaryHeaderRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J1:M1");
          let summaryContentRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J2:M2");
          summaryHeaderRow.values = summaryHeader;
          summaryContentRow.values = summaryContent;
          let formats = [[".000", ".000"]];
          summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).numberFormat = formats;
        }
      }

      // Translate an IP address into an integer.
      function ipAddressToInteger(ipAddress: string) {
        // Split the IP address into octets.
        let octets = ipAddress.split(".");

        // Create a number for each octet and do the math to create the integer value of the IP address.
        let fullNum =
          // Define an arbitrary number for the last octet.
          111 +
          parseInt(octets[2]) * 256 +
          parseInt(octets[1]) * 65536 +
          parseInt(octets[0]) * 16777216;
        return fullNum;
      }

      // Return the row number where the ip address is found.
      function findIP(ipLookupTable: number[][], n: number) {
        for (let i = 0; i < ipLookupTable.length; i++) {
          if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
            return i;
          }
        }
        return -1;
      }
    ```

5. <span data-ttu-id="92423-125">Переименуйте сценарий, чтобы **проанализировать загрузку веб-файлов** и сохранить его.</span><span class="sxs-lookup"><span data-stu-id="92423-125">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="92423-126">Выполнение скрипта</span><span class="sxs-lookup"><span data-stu-id="92423-126">Running the script</span></span>

<span data-ttu-id="92423-127">Перейдите к любому листу \*\*недели\* \*\* и запустите скрипт **анализа веб-загрузки** .</span><span class="sxs-lookup"><span data-stu-id="92423-127">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="92423-128">Сценарий применит условное форматирование и расположение лабеллинг к текущему листу.</span><span class="sxs-lookup"><span data-stu-id="92423-128">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="92423-129">Кроме того, будет обновлен лист **сводки** .</span><span class="sxs-lookup"><span data-stu-id="92423-129">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="92423-130">Перед выполнением скрипта</span><span class="sxs-lookup"><span data-stu-id="92423-130">Before running the script</span></span>

![Лист, отображающий необработанные данные веб-трафика.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="92423-132">После выполнения скрипта</span><span class="sxs-lookup"><span data-stu-id="92423-132">After running the script</span></span>

![Лист с отформатированными сведениями о расположении IP с предыдущими строками веб-трафика.](../../images/scenario-analyze-web-downloads-after.png)

![Сводная таблица и диаграмма, в которой перечисляются листы, на которых выполнен сценарий.](../../images/scenario-analyze-web-downloads-table.png)
