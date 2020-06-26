---
title: Интеграция сценариев Office с автоматизацией управления питанием
description: Как получить скрипты Office для Excel в Интернете, работая с рабочими процессами Power Автоматизация.
ms.date: 06/24/2020
localization_priority: Normal
ms.openlocfilehash: 977d9c88d75c8070eb729a443b4e8bc9a32e456d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878850"
---
# <a name="integrate-office-scripts-with-power-automate"></a><span data-ttu-id="315a4-103">Интеграция сценариев Office с автоматизацией управления питанием</span><span class="sxs-lookup"><span data-stu-id="315a4-103">Integrate Office Scripts with Power Automate</span></span>

<span data-ttu-id="315a4-104">[Power автоматизиру](https://flow.microsoft.com) интегрирует ваш сценарий в больший рабочий процесс.</span><span class="sxs-lookup"><span data-stu-id="315a4-104">[Power Automate](https://flow.microsoft.com) integrates your script into a larger workflow.</span></span> <span data-ttu-id="315a4-105">Вы можете использовать автоматизацию управления питанием, например добавить содержимое электронной почты в таблицу листа или создать действия в средствах управления проектами на основе комментариев к книгам.</span><span class="sxs-lookup"><span data-stu-id="315a4-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span> <span data-ttu-id="315a4-106">Если вы впервые используете автоматизированное управление питанием, рекомендуем [ознакомиться со статьей "начать автоматизацию](/power-automate/getting-started)".</span><span class="sxs-lookup"><span data-stu-id="315a4-106">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="315a4-107">Здесь вы можете узнать больше об автоматизации рабочих процессов для нескольких служб.</span><span class="sxs-lookup"><span data-stu-id="315a4-107">There, you can learn more about automating your workflows across multiple services.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="315a4-108">В настоящее время вы не можете запускать сценарии Office из [общего потока](/power-automate/share-buttons).</span><span class="sxs-lookup"><span data-stu-id="315a4-108">Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons).</span></span> <span data-ttu-id="315a4-109">Только пользователь, создавший сценарий, может запускать его, даже если вы автоматизируем Power.</span><span class="sxs-lookup"><span data-stu-id="315a4-109">Only the user who created a script can run it, even through Power Automate.</span></span>

## <a name="getting-started"></a><span data-ttu-id="315a4-110">Начало работы</span><span class="sxs-lookup"><span data-stu-id="315a4-110">Getting started</span></span>

<span data-ttu-id="315a4-111">Чтобы приступить к объединению сценариев Power автоматизированного и Office, следуйте рекомендациям, описанным в разделе [starting Scripts with Power Автоматизация](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="315a4-111">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="315a4-112">С его помощью вы узнаете, как создать последовательность, вызывающую простой сценарий.</span><span class="sxs-lookup"><span data-stu-id="315a4-112">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="315a4-113">После выполнения этого руководства и [автоматического запуска сценариев с помощью руководства Power Автоматизация](../tutorials/excel-power-automate-trigger.md) вернитесь сюда, чтобы узнать подробности об интеграции платформы.</span><span class="sxs-lookup"><span data-stu-id="315a4-113">After you've completed that tutorial and the [Automatically run scripts with Power Automate](../tutorials/excel-power-automate-trigger.md) tutorial, return here to learn details about the platform integrations.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="315a4-114">Соединитель Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="315a4-114">Excel Online (Business) connector</span></span>

<span data-ttu-id="315a4-115">[Соединители](/connectors/connectors) — это мосты между автоматизированной автоматизацией и приложениями.</span><span class="sxs-lookup"><span data-stu-id="315a4-115">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="315a4-116">[Соединитель Excel Online (Business)](/connectors/excelonlinebusiness) предоставляет потокам доступ к книгам Excel.</span><span class="sxs-lookup"><span data-stu-id="315a4-116">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="315a4-117">Действие "Запуск скрипта" позволяет вызывать любой сценарий Office, доступный через выбранную книгу.</span><span class="sxs-lookup"><span data-stu-id="315a4-117">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="315a4-118">Вы не можете выполнять сценарии с помощью потока, вы можете передавать данные в книгу и из нее с помощью скриптов.</span><span class="sxs-lookup"><span data-stu-id="315a4-118">Not only can you run scripts through a flow, you can pass data to and from the workbook with the flow through the scripts.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="315a4-119">Действие "Запуск скрипта" дает пользователям, использующим Microsoft Connector, значительный доступ к книге и ее данным.</span><span class="sxs-lookup"><span data-stu-id="315a4-119">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="315a4-120">Кроме того, существуют риски, связанные с безопасностью, с помощью скриптов, которые выполняют внешние вызовы API, как описано во [внешних вызовах от автоматизации Powering](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="315a4-120">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="315a4-121">Если администратор имеет дело с очень конфиденциальными данными, он может либо отключить Microsoft Excel Online Connector, либо ограничить доступ к сценариям Office с помощью [сценариев Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span><span class="sxs-lookup"><span data-stu-id="315a4-121">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="passing-data-from-power-automate-into-a-script"></a><span data-ttu-id="315a4-122">Передача данных из Power автоматизировать в сценарий</span><span class="sxs-lookup"><span data-stu-id="315a4-122">Passing data from Power Automate into a script</span></span>

<span data-ttu-id="315a4-123">Все входные данные сценария указываются как дополнительные параметры `main` функции.</span><span class="sxs-lookup"><span data-stu-id="315a4-123">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="315a4-124">Например, если вы хотите, чтобы сценарий принимал объект `string` , представляющий имя в качестве входных данных, вы можете изменить `main` подпись на `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="315a4-124">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="315a4-125">Когда вы настраиваете потоки в Power Автоматизация, вы можете указать входные данные скрипта в виде статических значений, [выражений](/power-automate/use-expressions-in-conditions)или динамического содержимого.</span><span class="sxs-lookup"><span data-stu-id="315a4-125">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="315a4-126">Подробные сведения о соединителе отдельных служб можно найти в [документации Power автоматизиру Connector](/connectors/).</span><span class="sxs-lookup"><span data-stu-id="315a4-126">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="315a4-127">При добавлении входных параметров в функцию сценария `main` учитывайте следующие ограничения и ограничения.</span><span class="sxs-lookup"><span data-stu-id="315a4-127">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="315a4-128">Первый параметр должен иметь тип `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="315a4-128">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="315a4-129">Имя параметра не имеет значения.</span><span class="sxs-lookup"><span data-stu-id="315a4-129">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="315a4-130">Каждый параметр должен иметь тип.</span><span class="sxs-lookup"><span data-stu-id="315a4-130">Every parameter must have a type.</span></span>

3. <span data-ttu-id="315a4-131">Основные типы,,,,,, `string` `number` `boolean` `any` `unknown` `object` и `undefined` поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="315a4-131">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="315a4-132">Массивы приведенных выше базовых типов поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="315a4-132">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="315a4-133">Вложенные массивы поддерживаются в качестве параметров (но не как типы возвращаемого значения).</span><span class="sxs-lookup"><span data-stu-id="315a4-133">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="315a4-134">Типы Union разрешены, если они являются объединением литералов, принадлежащих одному типу ( `string` , `number` или `boolean` ).</span><span class="sxs-lookup"><span data-stu-id="315a4-134">Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`).</span></span> <span data-ttu-id="315a4-135">Также поддерживаются объединения поддерживаемого типа с неопределенными.</span><span class="sxs-lookup"><span data-stu-id="315a4-135">Unions of a supported type with undefined are also supported.</span></span>

7. <span data-ttu-id="315a4-136">Типы объектов разрешены, если они содержат свойства типа `string` , `number` , `boolean` , поддерживаемых массивов или других поддерживаемых объектов.</span><span class="sxs-lookup"><span data-stu-id="315a4-136">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="315a4-137">В следующем примере показаны вложенные объекты, которые поддерживаются как типы параметров:</span><span class="sxs-lookup"><span data-stu-id="315a4-137">The following example shows nested objects that are supported as parameter types:</span></span>

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. <span data-ttu-id="315a4-138">Объекты должны иметь определение интерфейса или класса, определенное в сценарии.</span><span class="sxs-lookup"><span data-stu-id="315a4-138">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="315a4-139">Объект также может быть определен анонимно, как показано в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="315a4-139">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="315a4-140">Необязательные параметры разрешены и могут быть отмечены с помощью необязательного модификатора `?` (например, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="315a4-140">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="315a4-141">Допустимые значения параметров по умолчанию (например `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` ,.</span><span class="sxs-lookup"><span data-stu-id="315a4-141">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

## <a name="returning-data-from-a-script-back-to-power-automate"></a><span data-ttu-id="315a4-142">Возврат данных из скрипта в Power Автоматизация</span><span class="sxs-lookup"><span data-stu-id="315a4-142">Returning data from a script back to Power Automate</span></span>

<span data-ttu-id="315a4-143">Скрипты могут возвращать данные из книги для использования в качестве динамического контента в автоматизированном блоке управления питанием.</span><span class="sxs-lookup"><span data-stu-id="315a4-143">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="315a4-144">Как и в случае с входными параметрами, Автоматизация управления питанием применяет некоторые ограничения к типу возвращаемого значения.</span><span class="sxs-lookup"><span data-stu-id="315a4-144">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="315a4-145">Поддерживаются основные типы,,,, `string` `number` `boolean` `void` и `undefined` .</span><span class="sxs-lookup"><span data-stu-id="315a4-145">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="315a4-146">Типы объединения, используемые в качестве возвращаемых типов, действуют с теми же ограничениями, что и при использовании в качестве параметров сценария.</span><span class="sxs-lookup"><span data-stu-id="315a4-146">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="315a4-147">Типы массивов разрешены, если они имеют тип `string` , `number` или `boolean` .</span><span class="sxs-lookup"><span data-stu-id="315a4-147">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="315a4-148">Они также разрешены, если тип является поддерживаемым объединением или поддерживаемым типом литерала.</span><span class="sxs-lookup"><span data-stu-id="315a4-148">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="315a4-149">Типы объектов, используемые в качестве возвращаемых типов, действуют с теми же ограничениями, что и при использовании в качестве параметров сценария.</span><span class="sxs-lookup"><span data-stu-id="315a4-149">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="315a4-150">Неявная типизация поддерживается, несмотря на то, что они должны следовать тем же правилам, что и определенный тип.</span><span class="sxs-lookup"><span data-stu-id="315a4-150">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="315a4-151">Избегайте использования относительных ссылок</span><span class="sxs-lookup"><span data-stu-id="315a4-151">Avoid using relative references</span></span>

<span data-ttu-id="315a4-152">Power автоматизирует выполнение вашего сценария в выбранной книге Excel от вашего имени.</span><span class="sxs-lookup"><span data-stu-id="315a4-152">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="315a4-153">В этом случае книга может быть закрыта.</span><span class="sxs-lookup"><span data-stu-id="315a4-153">The workbook might be closed when this happens.</span></span> <span data-ttu-id="315a4-154">Любой API, зависящий от текущего состояния пользователя (например `Workbook.getActiveWorksheet` ,), не будет работать при использовании автоматизации Powering.</span><span class="sxs-lookup"><span data-stu-id="315a4-154">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="315a4-155">При проектировании скриптов обязательно используйте абсолютные ссылки на листы и диапазоны.</span><span class="sxs-lookup"><span data-stu-id="315a4-155">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="315a4-156">Приведенные ниже функции вызовут ошибку и завершатся ошибкой при вызове из скрипта в блоке автоматизации Power.</span><span class="sxs-lookup"><span data-stu-id="315a4-156">The following functions will throw an error and fail when called from a script in a Power Automate flow.</span></span>

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getActiveWorksheet`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`

## <a name="example"></a><span data-ttu-id="315a4-157">Пример</span><span class="sxs-lookup"><span data-stu-id="315a4-157">Example</span></span>

<span data-ttu-id="315a4-158">На следующем снимке экрана показан процесс автоматизации Power, который срабатывает при назначении вопроса [GitHub](https://github.com/) .</span><span class="sxs-lookup"><span data-stu-id="315a4-158">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="315a4-159">Поток выполняет сценарий, который добавляет ошибку в таблицу в книге Excel.</span><span class="sxs-lookup"><span data-stu-id="315a4-159">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="315a4-160">Если в этой таблице имеется пять или более проблем, посылается напоминание по электронной почте.</span><span class="sxs-lookup"><span data-stu-id="315a4-160">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![Пример процесса, показанный в редакторе автоматизации управления питанием.](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="315a4-162">`main`Функция скрипта ЗАДАЕТ идентификатор вопроса и заголовок вопроса в качестве входных параметров, а скрипт возвращает количество строк в таблице "ошибка".</span><span class="sxs-lookup"><span data-stu-id="315a4-162">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a><span data-ttu-id="315a4-163">См. также</span><span class="sxs-lookup"><span data-stu-id="315a4-163">See also</span></span>

- [<span data-ttu-id="315a4-164">Запуск сценариев Office в Excel в Интернете с помощью Power автоматизиру</span><span class="sxs-lookup"><span data-stu-id="315a4-164">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="315a4-165">Автоматический запуск сценариев с помощью автоматизации управления питанием</span><span class="sxs-lookup"><span data-stu-id="315a4-165">Automatically run scripts with Power Automate</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="315a4-166">Основы сценариев для сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="315a4-166">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="315a4-167">Начало работы с Power Automate</span><span class="sxs-lookup"><span data-stu-id="315a4-167">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="315a4-168">Справочная документация по Microsoft Online Connector (бизнес)</span><span class="sxs-lookup"><span data-stu-id="315a4-168">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
