---
title: Начало использования сценариев с помощью Power Automate
description: Руководство по использованию сценариев Office в Power Автоматизация через триггер, выполняемый вручную.
ms.date: 06/29/2020
localization_priority: Priority
ms.openlocfilehash: b2a10692929de90506f58851e9322afa63e15ca0
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043414"
---
# <a name="start-using-scripts-with-power-automate-preview"></a><span data-ttu-id="a4499-103">Начало работы со сценариями с помощью автоматизации управления (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="a4499-103">Start using scripts with Power Automate (preview)</span></span>

<span data-ttu-id="a4499-104">В этом руководстве рассказывается, как запускать скрипт Office для Excel в Интернете с помощью [автоматизации Powering](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="a4499-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a4499-105">Необходимые действия</span><span class="sxs-lookup"><span data-stu-id="a4499-105">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="a4499-106">В этом руководстве предполагается, что вы выполнили [запись, редактирование и создание сценариев Office в Excel в Интернете](excel-tutorial.md) .</span><span class="sxs-lookup"><span data-stu-id="a4499-106">This tutorial assumes you have completed the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="a4499-107">Подготовка книги</span><span class="sxs-lookup"><span data-stu-id="a4499-107">Prepare the workbook</span></span>

<span data-ttu-id="a4499-108">Автоматизация управления питанием не может использовать относительные ссылки, такие как `Workbook.getActiveWorksheet` доступ к компонентам книги.</span><span class="sxs-lookup"><span data-stu-id="a4499-108">Power Automate can't use relative references like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="a4499-109">Поэтому нам нужна книга и лист с согласованными именами, на которые может ссылаться Автоматизация управления питанием.</span><span class="sxs-lookup"><span data-stu-id="a4499-109">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="a4499-110">Создайте новую книгу с именем **миворкбук**.</span><span class="sxs-lookup"><span data-stu-id="a4499-110">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="a4499-111">В книге **миворкбук** Создайте лист под названием **туториалворкшит**.</span><span class="sxs-lookup"><span data-stu-id="a4499-111">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="a4499-112">Создание скрипта Office</span><span class="sxs-lookup"><span data-stu-id="a4499-112">Create an Office Script</span></span>

1. <span data-ttu-id="a4499-113">Перейдите на вкладку **Автоматизация** и выберите **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="a4499-113">Go to the **Automate** tab and select **Code Editor**.</span></span>

2. <span data-ttu-id="a4499-114">Выберите пункт **создать скрипт**.</span><span class="sxs-lookup"><span data-stu-id="a4499-114">Select **New Script**.</span></span>

3. <span data-ttu-id="a4499-115">Замените стандартный сценарий следующим.</span><span class="sxs-lookup"><span data-stu-id="a4499-115">Replace the default script with the following script.</span></span> <span data-ttu-id="a4499-116">Этот сценарий добавляет текущую дату и время в первые две ячейки листа **туториалворкшит** .</span><span class="sxs-lookup"><span data-stu-id="a4499-116">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

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

4. <span data-ttu-id="a4499-117">Переименуйте сценарий, чтобы **задать дату и время**.</span><span class="sxs-lookup"><span data-stu-id="a4499-117">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="a4499-118">Нажмите имя скрипта, чтобы изменить его.</span><span class="sxs-lookup"><span data-stu-id="a4499-118">Press the script name to change it.</span></span>

5. <span data-ttu-id="a4499-119">Сохраните скрипт, нажав **Сохранить скрипт**.</span><span class="sxs-lookup"><span data-stu-id="a4499-119">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="a4499-120">Создание автоматизированного рабочего процесса с помощью автоматизации управления питанием</span><span class="sxs-lookup"><span data-stu-id="a4499-120">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="a4499-121">Войдите на [сайт Power автоматизированного просмотра](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="a4499-121">Sign in to the [Power Automate preview site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="a4499-122">В меню, которое отображается в левой части экрана, нажмите кнопку **создать**.</span><span class="sxs-lookup"><span data-stu-id="a4499-122">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="a4499-123">В этом списке приводится список способов создания новых рабочих процессов.</span><span class="sxs-lookup"><span data-stu-id="a4499-123">This brings you to list of ways to create new workflows.</span></span>

    ![Кнопка "создать" в Power автоматизирует.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="a4499-125">В разделе **начать с пустого** раздела выберите **мгновенный ход**.</span><span class="sxs-lookup"><span data-stu-id="a4499-125">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="a4499-126">При этом создается рабочий процесс, активированный вручную.</span><span class="sxs-lookup"><span data-stu-id="a4499-126">This creates a manually activated workflow.</span></span>

    ![Вариант мгновенного потока для создания нового рабочего процесса.](../images/power-automate-tutorial-2.png)

4. <span data-ttu-id="a4499-128">В появившемся диалоговом окне введите имя для своего процесса в текстовом поле **имя процесса** , выберите **вручную запустить потоки** из списка **выберите способ запуска процесса**и нажмите кнопку **создать**.</span><span class="sxs-lookup"><span data-stu-id="a4499-128">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    ![Параметр ручного запуска для создания нового мгновенного движения.](../images/power-automate-tutorial-3.png)

5. <span data-ttu-id="a4499-130">Нажмите кнопку **создать шаг**.</span><span class="sxs-lookup"><span data-stu-id="a4499-130">Press **New step**.</span></span>

6. <span data-ttu-id="a4499-131">Перейдите на вкладку **Стандартная** и выберите **Excel Online (бизнес)**.</span><span class="sxs-lookup"><span data-stu-id="a4499-131">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Вариант Power Автоматизация для Excel Online (бизнес).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="a4499-133">В разделе **действия**выберите команду **выполнить скрипт (Предварительная версия)**.</span><span class="sxs-lookup"><span data-stu-id="a4499-133">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Параметр Power автоматизирует действие для сценария Run (Предварительная версия).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="a4499-135">Укажите следующие параметры для соединителя **сценария запуска** :</span><span class="sxs-lookup"><span data-stu-id="a4499-135">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="a4499-136">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="a4499-136">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="a4499-137">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="a4499-137">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="a4499-138">**Файл**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="a4499-138">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="a4499-139">**Сценарий**: Установка даты и времени</span><span class="sxs-lookup"><span data-stu-id="a4499-139">**Script**: Set date and time</span></span>

    ![Параметры соединителя для запуска скрипта в Power Автоматизация.](../images/power-automate-tutorial-6.png)

9. <span data-ttu-id="a4499-141">Нажмите кнопку **сохранить**.</span><span class="sxs-lookup"><span data-stu-id="a4499-141">Press **Save**.</span></span>

<span data-ttu-id="a4499-142">Теперь ваш поток готов к выполнению с помощью автоматизации управления питанием.</span><span class="sxs-lookup"><span data-stu-id="a4499-142">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="a4499-143">Вы можете протестировать его с помощью кнопки **тест** в редакторе потока или выполнить оставшиеся шаги руководства для запуска потока из коллекции потока.</span><span class="sxs-lookup"><span data-stu-id="a4499-143">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="a4499-144">Запуск скрипта с помощью Power автоматизиру</span><span class="sxs-lookup"><span data-stu-id="a4499-144">Run the script through Power Automate</span></span>

1. <span data-ttu-id="a4499-145">На главной странице Power Автоматизация выберите пункт **мои потоки**.</span><span class="sxs-lookup"><span data-stu-id="a4499-145">From the main Power Automate page, select **My flows**.</span></span>

    ![Кнопка "мои потоки" в Power автоматизирует.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="a4499-147">Выберите **мой поток обучения** в списке потоков, отображаемых на вкладке **мои потоки** . Здесь отображаются сведения о созданном ранее блоке.</span><span class="sxs-lookup"><span data-stu-id="a4499-147">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="a4499-148">Нажмите кнопку **выполнить**.</span><span class="sxs-lookup"><span data-stu-id="a4499-148">Press **Run**.</span></span>

    ![Кнопка "выполнить" в Power автоматизирует.](../images/power-automate-tutorial-8.png)

4. <span data-ttu-id="a4499-150">Откроется область задач для запуска процесса.</span><span class="sxs-lookup"><span data-stu-id="a4499-150">A task pane will appear for running the flow.</span></span> <span data-ttu-id="a4499-151">Если вам будет предложено **войти** в Excel Online, сделайте это, нажав кнопку **Continue (продолжить**).</span><span class="sxs-lookup"><span data-stu-id="a4499-151">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="a4499-152">Нажмите кнопку **поток выполнения**.</span><span class="sxs-lookup"><span data-stu-id="a4499-152">Press **Run flow**.</span></span> <span data-ttu-id="a4499-153">При этом выполняется поток, который запускает связанный сценарий Office.</span><span class="sxs-lookup"><span data-stu-id="a4499-153">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="a4499-154">Нажмите кнопку **done (Готово**).</span><span class="sxs-lookup"><span data-stu-id="a4499-154">Press **Done**.</span></span> <span data-ttu-id="a4499-155">Раздел **запуски** должен быть обновлен соответствующим образом.</span><span class="sxs-lookup"><span data-stu-id="a4499-155">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="a4499-156">Обновите страницу, чтобы увидеть результаты автоматизации Power.</span><span class="sxs-lookup"><span data-stu-id="a4499-156">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="a4499-157">Если она была успешно выполнена, перейдите в книгу, чтобы просмотреть обновленные ячейки.</span><span class="sxs-lookup"><span data-stu-id="a4499-157">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="a4499-158">Если это не удалось, проверьте параметры потока и запустите его еще раз.</span><span class="sxs-lookup"><span data-stu-id="a4499-158">If it failed, verify the flow's settings and run it a second time.</span></span>

    ![Автоматический выход Power, демонстрирующий успешный ход выполнения.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a><span data-ttu-id="a4499-160">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="a4499-160">Next steps</span></span>

<span data-ttu-id="a4499-161">Выполните [сценарии автоматического запуска с помощью руководства Power автоматизиру](excel-power-automate-trigger.md) .</span><span class="sxs-lookup"><span data-stu-id="a4499-161">Complete the [Automatically run scripts with Power Automate](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="a4499-162">В нем рассказывается, как передавать данные из службы рабочих процессов в сценарий Office.</span><span class="sxs-lookup"><span data-stu-id="a4499-162">It teaches you how to pass data from a workflow service to your Office Script.</span></span>
