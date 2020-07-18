---
title: Вызов сценариев из активированного вручную потока Power Automate
description: В этом руководстве рассказывается об использовании сценариев Office в Power Automate с помощью триггера с ручным срабатыванием.
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: 70fca2620973ecefe9eda40f02e28f064b713677
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160441"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a><span data-ttu-id="81958-103">Вызов сценариев из активированного вручную потока Power Automate (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="81958-103">Call scripts from a manual Power Automate flow (preview)</span></span>

<span data-ttu-id="81958-104">В этом руководстве объясняется, как запускать сценарий Office для Excel в Интернете с помощью [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="81958-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="81958-105">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="81958-105">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="81958-106">В этом руководстве предполагается, что вы прочитали руководство [Запись, изменение и создание сценариев Office для Excel в Интернете](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="81958-106">This tutorial assumes you have completed the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="81958-107">Подготовка книги</span><span class="sxs-lookup"><span data-stu-id="81958-107">Prepare the workbook</span></span>

<span data-ttu-id="81958-108">В Power Automate для доступа к компонентам книги нельзя использовать такие относительные ссылки, как `Workbook.getActiveWorksheet`.</span><span class="sxs-lookup"><span data-stu-id="81958-108">Power Automate can't use relative references like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="81958-109">Поэтому нужно использовать книгу и лист с именами, на которые может ссылаться Power Automate.</span><span class="sxs-lookup"><span data-stu-id="81958-109">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="81958-110">Создайте новую книгу под названием **MyWorkbook**.</span><span class="sxs-lookup"><span data-stu-id="81958-110">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="81958-111">В книге **MyWorkbook** создайте лист под названием **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="81958-111">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="81958-112">Создание сценария Office</span><span class="sxs-lookup"><span data-stu-id="81958-112">Create an Office Script</span></span>

1. <span data-ttu-id="81958-113">Откройте вкладку **Автоматизация** и запустите **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="81958-113">Go to the **Automate** tab and select **Code Editor**.</span></span>

2. <span data-ttu-id="81958-114">Выберите **Новый сценарий**.</span><span class="sxs-lookup"><span data-stu-id="81958-114">Select **New Script**.</span></span>

3. <span data-ttu-id="81958-115">Замените сценарий по умолчанию следующим сценарием.</span><span class="sxs-lookup"><span data-stu-id="81958-115">Replace the default script with the following script.</span></span> <span data-ttu-id="81958-116">Этот сценарий добавляет текущую дату и время в первые две ячейки листа **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="81958-116">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. <span data-ttu-id="81958-117">Переименуйте сценарий в **Установка даты и времени**.</span><span class="sxs-lookup"><span data-stu-id="81958-117">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="81958-118">Нажмите на имя сценария, чтобы изменить его.</span><span class="sxs-lookup"><span data-stu-id="81958-118">Press the script name to change it.</span></span>

5. <span data-ttu-id="81958-119">Сохраните сценарий, нажав кнопку **Сохранить сценарий**.</span><span class="sxs-lookup"><span data-stu-id="81958-119">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="81958-120">Создание автоматизированного рабочего процесса с помощью Power Automate</span><span class="sxs-lookup"><span data-stu-id="81958-120">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="81958-121">Войдите на [сайт Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="81958-121">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="81958-122">В меню, которое отображается в левой части экрана, нажмите клавишу **Создать**.</span><span class="sxs-lookup"><span data-stu-id="81958-122">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="81958-123">Откроется список способов создания новых рабочих процессов.</span><span class="sxs-lookup"><span data-stu-id="81958-123">This brings you to list of ways to create new workflows.</span></span>

    ![Кнопка "Создать" в Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="81958-125">В разделе **Создание нового** выберите пункт **Мгновенный поток**.</span><span class="sxs-lookup"><span data-stu-id="81958-125">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="81958-126">В результате будет создан активированный вручную рабочий процесс.</span><span class="sxs-lookup"><span data-stu-id="81958-126">This creates a manually activated workflow.</span></span>

    ![Способ мгновенного потока для создания нового рабочего процесса.](../images/power-automate-tutorial-2.png)

4. <span data-ttu-id="81958-128">В открывшемся диалоговом окне введите имя для своего потока в поле **Имя потока**, выберите **Запустить поток вручную** из списка вариантов в разделе **Выбор способа запуска потока**и нажмите **Создать**.</span><span class="sxs-lookup"><span data-stu-id="81958-128">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    ![Способ запуска потока вручную для создания нового мгновенного потока.](../images/power-automate-tutorial-3.png)

    <span data-ttu-id="81958-130">Обратите внимание: запускаемый вручную поток — это лишь один из многих типов потоков.</span><span class="sxs-lookup"><span data-stu-id="81958-130">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="81958-131">В следующем руководстве описывается создание потока, который будет выполняться автоматически при получении вами сообщения электронной почты.</span><span class="sxs-lookup"><span data-stu-id="81958-131">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="81958-132">Нажмите клавишу **Следующий шаг**.</span><span class="sxs-lookup"><span data-stu-id="81958-132">Press **New step**.</span></span>

6. <span data-ttu-id="81958-133">Откройте вкладку **Стандартные**, а затем выберите **Excel Online (бизнес)**.</span><span class="sxs-lookup"><span data-stu-id="81958-133">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Функция Power Automate для Excel Online (бизнес).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="81958-135">В разделе **Действия** выберите **Запустить сценарий (предварительный просмотр)**.</span><span class="sxs-lookup"><span data-stu-id="81958-135">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Вариант действия Power Automate "Запуск сценария" (предварительный просмотр).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="81958-137">Определите указанные ниже параметры для соединителя **Запуск сценария**.</span><span class="sxs-lookup"><span data-stu-id="81958-137">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="81958-138">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="81958-138">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="81958-139">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="81958-139">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="81958-140">**Файл**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="81958-140">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="81958-141">**Сценарий**: Установка даты и времени</span><span class="sxs-lookup"><span data-stu-id="81958-141">**Script**: Set date and time</span></span>

    ![Параметры соединителя для запуска сценария в Power Automate.](../images/power-automate-tutorial-6.png)

9. <span data-ttu-id="81958-143">Нажмите **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="81958-143">Press **Save**.</span></span>

<span data-ttu-id="81958-144">Теперь ваш поток готов к запуску с помощью Power Automate.</span><span class="sxs-lookup"><span data-stu-id="81958-144">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="81958-145">Вы можете проверить его с помощью кнопки **Тест** в редакторе потока или выполнить остальные действия согласно руководству, чтобы запустить поток из вашей коллекции потоков.</span><span class="sxs-lookup"><span data-stu-id="81958-145">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="81958-146">Запуск сценария с помощью Power Automate</span><span class="sxs-lookup"><span data-stu-id="81958-146">Run the script through Power Automate</span></span>

1. <span data-ttu-id="81958-147">На главной странице Power Automate выберите **Мои потоки**.</span><span class="sxs-lookup"><span data-stu-id="81958-147">From the main Power Automate page, select **My flows**.</span></span>

    ![Кнопка "Мои потоки" в Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="81958-149">Выберите **Мой учебный поток** из списка во вкладке **Мои потоки**. При этом будут показаны подробные сведения о потоке, который мы создали ранее.</span><span class="sxs-lookup"><span data-stu-id="81958-149">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="81958-150">Нажмите кнопку **Запуск**.</span><span class="sxs-lookup"><span data-stu-id="81958-150">Press **Run**.</span></span>

    ![Кнопка "Запуск" в Power Automate.](../images/power-automate-tutorial-8.png)

4. <span data-ttu-id="81958-152">Появится панель задач для запуска потока.</span><span class="sxs-lookup"><span data-stu-id="81958-152">A task pane will appear for running the flow.</span></span> <span data-ttu-id="81958-153">Когда будет предложено выполнить **Вход** в Excel Online, нажмите кнопку **Продолжить**.</span><span class="sxs-lookup"><span data-stu-id="81958-153">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="81958-154">Щелкните **Запустить поток**.</span><span class="sxs-lookup"><span data-stu-id="81958-154">Press **Run flow**.</span></span> <span data-ttu-id="81958-155">При этом запустится поток, выполняющий связанный сценарий Office.</span><span class="sxs-lookup"><span data-stu-id="81958-155">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="81958-156">Нажмите кнопку **Готово**.</span><span class="sxs-lookup"><span data-stu-id="81958-156">Press **Done**.</span></span> <span data-ttu-id="81958-157">Вы можете заметить, что раздел **Запуски** соответствующим образом обновлен.</span><span class="sxs-lookup"><span data-stu-id="81958-157">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="81958-158">Обновите страницу, чтобы увидеть результаты работы Power Automate.</span><span class="sxs-lookup"><span data-stu-id="81958-158">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="81958-159">После этого перейдите в книгу, где должны отобразиться обновленные ячейки.</span><span class="sxs-lookup"><span data-stu-id="81958-159">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="81958-160">В случае неудачи проверьте параметры этого потока и запустите его еще раз.</span><span class="sxs-lookup"><span data-stu-id="81958-160">If it failed, verify the flow's settings and run it a second time.</span></span>

    ![В результатах работы Power Automate показано успешное выполнение потока.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a><span data-ttu-id="81958-162">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="81958-162">Next steps</span></span>

<span data-ttu-id="81958-163">Прочитайте раздел руководства [Передача данных сценариям в автоматически запускаемом потоке Power Automate](excel-power-automate-trigger.md).</span><span class="sxs-lookup"><span data-stu-id="81958-163">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="81958-164">В нем рассказывается о том, как передать данные из службы рабочего процесса в ваш сценарий Office и запустить поток Power Automate при возникновении определенных событий.</span><span class="sxs-lookup"><span data-stu-id="81958-164">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
