---
title: Автоматический запуск сценариев с помощью Power Automate
description: Руководство по запуску сценариев Office для Excel в Интернете с помощью автоматизированного внешнего триггера (получение почты через Outlook).
ms.date: 06/29/2020
localization_priority: Priority
ms.openlocfilehash: a750197d6b5ae770ad7d2e17b3ee00dc65ee8875
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043421"
---
# <a name="automatically-run-scripts-with-power-automate-preview"></a><span data-ttu-id="3d93a-103">Автоматический запуск сценариев с помощью автоматизации управления (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="3d93a-103">Automatically run scripts with Power Automate (preview)</span></span>

<span data-ttu-id="3d93a-104">В этом руководстве рассказывается, как использовать скрипт Office для Excel в Интернете с автоматизированным рабочим процессом [Power автоматизировать](https://flow.microsoft.com) .</span><span class="sxs-lookup"><span data-stu-id="3d93a-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="3d93a-105">Ваш сценарий будет автоматически запускаться при каждом получении сообщения электронной почты, записи данных из электронной почты в книгу Excel.</span><span class="sxs-lookup"><span data-stu-id="3d93a-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3d93a-106">Необходимые действия</span><span class="sxs-lookup"><span data-stu-id="3d93a-106">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="3d93a-107">В этом руководстве предполагается, что вы завершили [работу со сценариями Office в Excel в Интернете с помощью руководства Power автоматизиру](excel-power-automate-manual.md) .</span><span class="sxs-lookup"><span data-stu-id="3d93a-107">This tutorial assumes you have completed the [Run Office Scripts in Excel on the web with Power Automate](excel-power-automate-manual.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="3d93a-108">Подготовка книги</span><span class="sxs-lookup"><span data-stu-id="3d93a-108">Prepare the workbook</span></span>

<span data-ttu-id="3d93a-109">Автоматизация управления питанием не может использовать [относительные ссылки](../develop/power-automate-integration.md#avoid-using-relative-references) , такие как `Workbook.getActiveWorksheet` доступ к компонентам книги.</span><span class="sxs-lookup"><span data-stu-id="3d93a-109">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="3d93a-110">Поэтому нам потребуется книга и лист с согласованными именами для автоматизации управления питанием.</span><span class="sxs-lookup"><span data-stu-id="3d93a-110">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="3d93a-111">Создайте новую книгу с именем **миворкбук**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-111">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="3d93a-112">Перейдите на вкладку **Автоматизация** и выберите **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-112">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="3d93a-113">Выберите пункт **создать скрипт**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-113">Select **New Script**.</span></span>

4. <span data-ttu-id="3d93a-114">Замените существующий код приведенным ниже скриптом и нажмите кнопку **выполнить**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-114">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="3d93a-115">Это приведет к настройке книги с одинаковыми именами листа, таблицы и сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="3d93a-115">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Add a new worksheet to store our email table
      let emailsSheet = workbook.addWorksheet("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").setValues([
        ["Date", "Day of the week", "Email address", "Subject"]
      ]);
      let newTable = workbook.addTable(emailsSheet.getRange("A1:D2"), true);
      newTable.setName("EmailTable");

      // Add a new PivotTable to a new worksheet
      let pivotWorksheet = workbook.addWorksheet("SubjectPivot");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script-for-your-automated-workflow"></a><span data-ttu-id="3d93a-116">Создание скрипта Office для автоматического рабочего процесса</span><span class="sxs-lookup"><span data-stu-id="3d93a-116">Create an Office Script for your automated workflow</span></span>

<span data-ttu-id="3d93a-117">Давайте создадим сценарий, который записывает данные из сообщения электронной почты.</span><span class="sxs-lookup"><span data-stu-id="3d93a-117">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="3d93a-118">Мы хотим узнать, каким из дней недели мы получаем самую большую часть почты и сколько уникальных отправители отправляют эту почту.</span><span class="sxs-lookup"><span data-stu-id="3d93a-118">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="3d93a-119">Книга содержит таблицу с **датой**, **днем недели**, **адресом электронной почты**и столбцами **subject** .</span><span class="sxs-lookup"><span data-stu-id="3d93a-119">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="3d93a-120">На листе также имеется сводная таблица, сводная в **день недели** и **адрес электронной почты** (иерархии строк).</span><span class="sxs-lookup"><span data-stu-id="3d93a-120">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="3d93a-121">Количество уникальных **субъектов** — это сводные сведения, отображаемые (иерархия данных).</span><span class="sxs-lookup"><span data-stu-id="3d93a-121">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="3d93a-122">После обновления таблицы электронной почты мы будем обновлять сценарий в виде сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="3d93a-122">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="3d93a-123">В **редакторе кода**выберите **создать скрипт**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-123">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="3d93a-124">В процессе, который мы создадим в руководстве ниже, будут отправляться сведения о скрипте для каждого полученного сообщения.</span><span class="sxs-lookup"><span data-stu-id="3d93a-124">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="3d93a-125">Сценарий должен принимать эти данные с помощью параметров в `main` функции.</span><span class="sxs-lookup"><span data-stu-id="3d93a-125">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="3d93a-126">Замените стандартный сценарий следующим:</span><span class="sxs-lookup"><span data-stu-id="3d93a-126">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="3d93a-127">Скрипту требуется доступ к таблице и сводной таблице книги.</span><span class="sxs-lookup"><span data-stu-id="3d93a-127">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="3d93a-128">Добавьте следующий код в текст скрипта после открытия `{` :</span><span class="sxs-lookup"><span data-stu-id="3d93a-128">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="3d93a-129">`dateReceived`Параметр имеет тип `string` .</span><span class="sxs-lookup"><span data-stu-id="3d93a-129">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="3d93a-130">Давайте преобразуйте его в [ `Date` объект](../develop/javascript-objects.md#date) , чтобы мы могли легко получить день недели.</span><span class="sxs-lookup"><span data-stu-id="3d93a-130">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="3d93a-131">После этого необходимо сопоставить значение числа дней недели с более удобочитаемой версией.</span><span class="sxs-lookup"><span data-stu-id="3d93a-131">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="3d93a-132">Добавьте следующий код в конец скрипта перед закрытием `}` :</span><span class="sxs-lookup"><span data-stu-id="3d93a-132">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Parse the received date string.
    let date = new Date(dateReceived);

    // Convert number representing the day of the week into the name of the day.
    let dayText : string;
    switch (date.getDay()) {
      case 0:
        dayText = "Sunday";
        break;
      case 1:
        dayText = "Monday";
        break;
      case 2:
        dayText = "Tuesday";
        break;
      case 3:
        dayText = "Wednesday";
        break;
      case 4:
        dayText = "Thursday";
        break;
      case 5:
        dayText = "Friday";
        break;
      default:
        dayText = "Saturday";
        break;
    }
    ```

5. <span data-ttu-id="3d93a-133">`subject`Строка может содержать тег ответа "Re:".</span><span class="sxs-lookup"><span data-stu-id="3d93a-133">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="3d93a-134">Удалим это значение из строки, чтобы сообщения электронной почты в одном и том же потоке совпадали с темой таблицы.</span><span class="sxs-lookup"><span data-stu-id="3d93a-134">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="3d93a-135">Добавьте следующий код в конец скрипта перед закрытием `}` :</span><span class="sxs-lookup"><span data-stu-id="3d93a-135">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="3d93a-136">Теперь, когда данные электронной почты были отформатированы по своему вкусу, добавим строку в таблицу электронной почты.</span><span class="sxs-lookup"><span data-stu-id="3d93a-136">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="3d93a-137">Добавьте следующий код в конец скрипта перед закрытием `}` :</span><span class="sxs-lookup"><span data-stu-id="3d93a-137">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. <span data-ttu-id="3d93a-138">Теперь убедитесь, что сводная таблица обновлена.</span><span class="sxs-lookup"><span data-stu-id="3d93a-138">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="3d93a-139">Добавьте следующий код в конец скрипта перед закрытием `}` :</span><span class="sxs-lookup"><span data-stu-id="3d93a-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="3d93a-140">Переименуйте **запись** в сценарии и нажмите кнопку **Сохранить скрипт**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-140">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="3d93a-141">Теперь ваш сценарий готов к работе с рабочими процессами Power автоматизировать.</span><span class="sxs-lookup"><span data-stu-id="3d93a-141">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="3d93a-142">Он должен выглядеть так, как показано в следующем сценарии:</span><span class="sxs-lookup"><span data-stu-id="3d93a-142">It should look like the following script:</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  from: string,
  dateReceived: string,
  subject: string) {
  // Get the email table.
  let emailWorksheet = workbook.getWorksheet("Emails");
  let table = emailWorksheet.getTable("EmailTable");

  // Get the PivotTable.
  let pivotTableWorksheet = workbook.getWorksheet("Pivot");
  let pivotTable = pivotTableWorksheet.getPivotTable("SubjectPivot");

  // Parse the received date string.
  let date = new Date(dateReceived);

  // Convert number representing the day of the week into the name of the day.
  let dayText: string;
  switch (date.getDay()) {
    case 0:
      dayText = "Sunday";
      break;
    case 1:
      dayText = "Monday";
      break;
    case 2:
      dayText = "Tuesday";
      break;
    case 3:
      dayText = "Wednesday";
      break;
    case 4:
      dayText = "Thursday";
      break;
    case 5:
      dayText = "Friday";
      break;
    default:
      dayText = "Saturday";
      break;
  }

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayText, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="3d93a-143">Создание автоматизированного рабочего процесса с помощью автоматизации управления питанием</span><span class="sxs-lookup"><span data-stu-id="3d93a-143">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="3d93a-144">Войдите на [сайт Power автоматизированного просмотра](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="3d93a-144">Sign in to the [Power Automate preview site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="3d93a-145">В меню, которое отображается в левой части экрана, нажмите кнопку **создать**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-145">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="3d93a-146">В этом списке приводится список способов создания новых рабочих процессов.</span><span class="sxs-lookup"><span data-stu-id="3d93a-146">This brings you to list of ways to create new workflows.</span></span>

    ![Кнопка "создать" в Power автоматизирует.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="3d93a-148">В разделе **начать с пустого** раздела выберите **автоматизированный процесс**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-148">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="3d93a-149">При этом создается рабочий процесс, запущенный событием, например получение сообщения электронной почты.</span><span class="sxs-lookup"><span data-stu-id="3d93a-149">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![Параметр автоматического перенаправления в Power автоматизируется.](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="3d93a-151">В появившемся диалоговом окне введите имя для своего процесса в текстовом поле **имя процесса** .</span><span class="sxs-lookup"><span data-stu-id="3d93a-151">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="3d93a-152">Затем выберите **, когда новое сообщение поступает** из списка вариантов в разделе **Выберите триггер вашего процесса**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-152">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="3d93a-153">Вам может потребоваться выполнить поиск параметра с помощью поля поиска.</span><span class="sxs-lookup"><span data-stu-id="3d93a-153">You may need to search for the option using the search box.</span></span> <span data-ttu-id="3d93a-154">Наконец, нажмите кнопку **создать**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-154">Finally, press **Create**.</span></span>

    ![Часть создание автоматизированного окна "блок" в Power автоматизиру, в котором отображается параметр "Новое поступление почты".](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="3d93a-156">В этом руководстве используется Outlook.</span><span class="sxs-lookup"><span data-stu-id="3d93a-156">This tutorial uses Outlook.</span></span> <span data-ttu-id="3d93a-157">Вместо этого вы можете использовать предпочтительную почтовую службу, хотя некоторые варианты могут отличаться.</span><span class="sxs-lookup"><span data-stu-id="3d93a-157">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="3d93a-158">Нажмите кнопку **создать шаг**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-158">Press **New step**.</span></span>

6. <span data-ttu-id="3d93a-159">Перейдите на вкладку **Стандартная** и выберите **Excel Online (бизнес)**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-159">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Вариант Power Автоматизация для Excel Online (бизнес).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="3d93a-161">В разделе **действия**выберите команду **выполнить скрипт (Предварительная версия)**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-161">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Параметр Power автоматизирует действие для сценария Run (Предварительная версия).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="3d93a-163">Укажите следующие параметры для соединителя **сценария запуска** :</span><span class="sxs-lookup"><span data-stu-id="3d93a-163">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="3d93a-164">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="3d93a-164">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="3d93a-165">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="3d93a-165">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="3d93a-166">**Файл**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="3d93a-166">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="3d93a-167">**Сценарий**: запись электронной почты</span><span class="sxs-lookup"><span data-stu-id="3d93a-167">**Script**: Record Email</span></span>
    - <span data-ttu-id="3d93a-168">**от**: from *(динамическое содержимое из Outlook)*</span><span class="sxs-lookup"><span data-stu-id="3d93a-168">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="3d93a-169">**датерецеивед**: время получения *(динамическое содержимое из Outlook)*</span><span class="sxs-lookup"><span data-stu-id="3d93a-169">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="3d93a-170">**Тема**: subject *(динамическое содержимое из Outlook)*</span><span class="sxs-lookup"><span data-stu-id="3d93a-170">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="3d93a-171">*Обратите внимание, что параметры для сценария будут отображаться только после выбора сценария.*</span><span class="sxs-lookup"><span data-stu-id="3d93a-171">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![Параметр Power автоматизирует действие для сценария Run (Предварительная версия).](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="3d93a-173">Нажмите кнопку **сохранить**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-173">Press **Save**.</span></span>

<span data-ttu-id="3d93a-174">Теперь ваш процесс включен.</span><span class="sxs-lookup"><span data-stu-id="3d93a-174">Your flow is now enabled.</span></span> <span data-ttu-id="3d93a-175">При каждом получении электронного сообщения через Outlook он будет автоматически запускаться.</span><span class="sxs-lookup"><span data-stu-id="3d93a-175">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="3d93a-176">Управление сценарием в Power Автоматизация</span><span class="sxs-lookup"><span data-stu-id="3d93a-176">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="3d93a-177">На главной странице Power Автоматизация выберите пункт **мои потоки**.</span><span class="sxs-lookup"><span data-stu-id="3d93a-177">From the main Power Automate page, select **My flows**.</span></span>

    ![Кнопка "мои потоки" в Power автоматизирует.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="3d93a-179">Выберите свой ход.</span><span class="sxs-lookup"><span data-stu-id="3d93a-179">Select your flow.</span></span> <span data-ttu-id="3d93a-180">Здесь вы можете просмотреть историю запуска.</span><span class="sxs-lookup"><span data-stu-id="3d93a-180">Here you can see the run history.</span></span> <span data-ttu-id="3d93a-181">Вы можете обновить страницу или нажать кнопку обновить **все запуски** , чтобы обновить журнал.</span><span class="sxs-lookup"><span data-stu-id="3d93a-181">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="3d93a-182">Процесс будет запущен вскоре после получения сообщения электронной почты.</span><span class="sxs-lookup"><span data-stu-id="3d93a-182">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="3d93a-183">Протестируйте процесс отправки почты.</span><span class="sxs-lookup"><span data-stu-id="3d93a-183">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="3d93a-184">При активации этого процесса и успешном выполнении сценария должна отобразиться таблица книги и обновление сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="3d93a-184">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![Таблица электронной почты после потока выполняется несколько раз.](../images/power-automate-params-tutorial-4.png)

![Сводная таблица после выполнения потока в несколько раз.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="3d93a-187">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="3d93a-187">Next steps</span></span>

<span data-ttu-id="3d93a-188">Для получения дополнительных сведений о подключении сценариев Office к автоматизации управления питанием посетите страницу [Запуск сценариев Office с помощью Power автоматизировать](../develop/power-automate-integration.md) .</span><span class="sxs-lookup"><span data-stu-id="3d93a-188">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="3d93a-189">Вы также можете ознакомиться с [примером сценария автоматизированных напоминаний](../resources/scenarios/task-reminders.md) о том, как объединять сценарии Office и автоматизацию управления питанием с помощью адаптивных карточек Teams.</span><span class="sxs-lookup"><span data-stu-id="3d93a-189">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
