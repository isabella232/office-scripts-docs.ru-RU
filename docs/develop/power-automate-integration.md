---
title: Запуск сценариев Office с помощью автоматизации управления питанием
description: Как получить скрипты Office для Excel в Интернете, работая с рабочими процессами Power Автоматизация.
ms.date: 07/24/2020
localization_priority: Normal
ms.openlocfilehash: 87bd4e15ef7680a7456077494e3fda8208d6b9d8
ms.sourcegitcommit: e9a8ef5f56177ea9a3d2fc5ac636368e5bdae1f4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/01/2020
ms.locfileid: "47321574"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="a9f81-103">Запуск сценариев Office с помощью автоматизации управления питанием</span><span class="sxs-lookup"><span data-stu-id="a9f81-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="a9f81-104">[Power автоматизиру](https://flow.microsoft.com) позволяет добавлять скрипты Office в более крупный автоматизированный рабочий процесс.</span><span class="sxs-lookup"><span data-stu-id="a9f81-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="a9f81-105">Вы можете использовать автоматизацию управления питанием, например добавить содержимое электронной почты в таблицу листа или создать действия в средствах управления проектами на основе комментариев к книгам.</span><span class="sxs-lookup"><span data-stu-id="a9f81-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="getting-started"></a><span data-ttu-id="a9f81-106">Начало работы</span><span class="sxs-lookup"><span data-stu-id="a9f81-106">Getting started</span></span>

<span data-ttu-id="a9f81-107">Если вы впервые используете автоматизированное управление питанием, рекомендуем [ознакомиться со статьей "начать автоматизацию](/power-automate/getting-started)".</span><span class="sxs-lookup"><span data-stu-id="a9f81-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="a9f81-108">Здесь вы можете узнать больше о всех возможных возможностях автоматизации.</span><span class="sxs-lookup"><span data-stu-id="a9f81-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="a9f81-109">В этом разделе приведены сведения о том, как сценарии Office работают с автоматизацией управления питанием и как они могут помочь повысить удобство работы с Excel.</span><span class="sxs-lookup"><span data-stu-id="a9f81-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="a9f81-110">Чтобы приступить к объединению сценариев Power автоматизированного и Office, следуйте рекомендациям, описанным в разделе [starting Scripts with Power Автоматизация](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="a9f81-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="a9f81-111">С его помощью вы узнаете, как создать последовательность, вызывающую простой сценарий.</span><span class="sxs-lookup"><span data-stu-id="a9f81-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="a9f81-112">После выполнения этого руководства и [передачи данных сценариям в руководстве автоматизированного управления энергопотреблением](../tutorials/excel-power-automate-trigger.md) вернитесь сюда для получения подробных сведений о подключении сценариев Office к автоматизации потоков Power автоматизированного.</span><span class="sxs-lookup"><span data-stu-id="a9f81-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="a9f81-113">Соединитель Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="a9f81-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="a9f81-114">[Соединители](/connectors/connectors) — это мосты между автоматизированной автоматизацией и приложениями.</span><span class="sxs-lookup"><span data-stu-id="a9f81-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="a9f81-115">[Соединитель Excel Online (Business)](/connectors/excelonlinebusiness) предоставляет потокам доступ к книгам Excel.</span><span class="sxs-lookup"><span data-stu-id="a9f81-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="a9f81-116">Действие "Запуск скрипта" позволяет вызывать любой сценарий Office, доступный через выбранную книгу.</span><span class="sxs-lookup"><span data-stu-id="a9f81-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="a9f81-117">Вы также можете предоставить входные параметры скриптов, чтобы данные могли быть предоставлены с помощью этого процесса, или чтобы скрипт возвращал сведения для последующих шагов в этом блоке.</span><span class="sxs-lookup"><span data-stu-id="a9f81-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a9f81-118">Действие "Запуск скрипта" дает пользователям, использующим Microsoft Connector, значительный доступ к книге и ее данным.</span><span class="sxs-lookup"><span data-stu-id="a9f81-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="a9f81-119">Кроме того, существуют риски, связанные с безопасностью, с помощью скриптов, которые выполняют внешние вызовы API, как описано во [внешних вызовах от автоматизации Powering](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="a9f81-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="a9f81-120">Если администратор имеет дело с очень конфиденциальными данными, он может либо отключить Microsoft Excel Online Connector, либо ограничить доступ к сценариям Office с помощью [сценариев Office](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="a9f81-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="a9f81-121">Передача данных в потоках для сценариев</span><span class="sxs-lookup"><span data-stu-id="a9f81-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="a9f81-122">Power автоматизиру позволяет передавать фрагменты данных между действиями в рамках своего процесса.</span><span class="sxs-lookup"><span data-stu-id="a9f81-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="a9f81-123">Сценарии можно настроить так, чтобы они принимали нужные вам типы сведений и возвращать что угодно из вашей книги, которую вы хотите использовать в вашем движении.</span><span class="sxs-lookup"><span data-stu-id="a9f81-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="a9f81-124">Входные данные для вашего скрипта задаются путем добавления параметров в `main` функцию (в дополнение к `workbook: ExcelScript.Workbook` ).</span><span class="sxs-lookup"><span data-stu-id="a9f81-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="a9f81-125">Выходные данные сценария объявляются путем добавления типа возвращаемого значения в `main` .</span><span class="sxs-lookup"><span data-stu-id="a9f81-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="a9f81-126">При создании блока "выполнить скрипт" в потоке заполняются допустимые параметры и возвращаемые типы.</span><span class="sxs-lookup"><span data-stu-id="a9f81-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="a9f81-127">Если вы изменяете параметры или типы возвращаемых данных в вашем сценарии, вам потребуется повторить блок потока "Run script".</span><span class="sxs-lookup"><span data-stu-id="a9f81-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="a9f81-128">Это гарантирует, что данные анализируются правильно.</span><span class="sxs-lookup"><span data-stu-id="a9f81-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="a9f81-129">В следующих разделах рассматриваются входные и выходные данные для сценариев, используемых в автоматизации Powering.</span><span class="sxs-lookup"><span data-stu-id="a9f81-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="a9f81-130">Если вы хотите получить практический подход к освоению этой статьи, ознакомьтесь со статьей " [Передача данных в скрипты в руководстве по потоку автоматизированного управления питанием](../tutorials/excel-power-automate-trigger.md) " или изучите пример сценария [автоматизированной задачи "напоминания](../resources/scenarios/task-reminders.md) ".</span><span class="sxs-lookup"><span data-stu-id="a9f81-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="a9f81-131">`main` Параметры: передача данных в скрипт</span><span class="sxs-lookup"><span data-stu-id="a9f81-131">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="a9f81-132">Все входные данные сценария указываются как дополнительные параметры `main` функции.</span><span class="sxs-lookup"><span data-stu-id="a9f81-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="a9f81-133">Например, если вы хотите, чтобы сценарий принимал объект `string` , представляющий имя в качестве входных данных, вы можете изменить `main` подпись на `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="a9f81-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="a9f81-134">Когда вы настраиваете потоки в Power Автоматизация, вы можете указать входные данные скрипта в виде статических значений, [выражений](/power-automate/use-expressions-in-conditions)или динамического содержимого.</span><span class="sxs-lookup"><span data-stu-id="a9f81-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="a9f81-135">Подробные сведения о соединителе отдельных служб можно найти в [документации Power автоматизиру Connector](/connectors/).</span><span class="sxs-lookup"><span data-stu-id="a9f81-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="a9f81-136">При добавлении входных параметров в функцию сценария `main` учитывайте следующие ограничения и ограничения.</span><span class="sxs-lookup"><span data-stu-id="a9f81-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="a9f81-137">Первый параметр должен иметь тип `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="a9f81-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="a9f81-138">Имя параметра не имеет значения.</span><span class="sxs-lookup"><span data-stu-id="a9f81-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="a9f81-139">Каждый параметр должен иметь тип (например, `string` или `number` ).</span><span class="sxs-lookup"><span data-stu-id="a9f81-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="a9f81-140">Основные типы,,,,,, `string` `number` `boolean` `any` `unknown` `object` и `undefined` поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="a9f81-140">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="a9f81-141">Массивы приведенных выше базовых типов поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="a9f81-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="a9f81-142">Вложенные массивы поддерживаются в качестве параметров (но не как типы возвращаемого значения).</span><span class="sxs-lookup"><span data-stu-id="a9f81-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="a9f81-143">Типы Union разрешены, если они являются объединением литералов, принадлежащих одному типу (например, `"Left" | "Right"` ).</span><span class="sxs-lookup"><span data-stu-id="a9f81-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="a9f81-144">Также поддерживаются объединения поддерживаемого типа с неопределенной версией (например, `string | undefined` ).</span><span class="sxs-lookup"><span data-stu-id="a9f81-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="a9f81-145">Типы объектов разрешены, если они содержат свойства типа `string` , `number` , `boolean` , поддерживаемых массивов или других поддерживаемых объектов.</span><span class="sxs-lookup"><span data-stu-id="a9f81-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="a9f81-146">В следующем примере показаны вложенные объекты, которые поддерживаются как типы параметров:</span><span class="sxs-lookup"><span data-stu-id="a9f81-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="a9f81-147">Объекты должны иметь определение интерфейса или класса, определенное в сценарии.</span><span class="sxs-lookup"><span data-stu-id="a9f81-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="a9f81-148">Объект также может быть определен анонимно, как показано в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="a9f81-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="a9f81-149">Необязательные параметры разрешены и могут быть отмечены с помощью необязательного модификатора `?` (например, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="a9f81-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="a9f81-150">Допустимые значения параметров по умолчанию (например `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` ,.</span><span class="sxs-lookup"><span data-stu-id="a9f81-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="returning-data-from-a-script"></a><span data-ttu-id="a9f81-151">Возвращение данных из скрипта</span><span class="sxs-lookup"><span data-stu-id="a9f81-151">Returning data from a script</span></span>

<span data-ttu-id="a9f81-152">Скрипты могут возвращать данные из книги для использования в качестве динамического контента в автоматизированном блоке управления питанием.</span><span class="sxs-lookup"><span data-stu-id="a9f81-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="a9f81-153">Как и в случае с входными параметрами, Автоматизация управления питанием применяет некоторые ограничения к типу возвращаемого значения.</span><span class="sxs-lookup"><span data-stu-id="a9f81-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="a9f81-154">Поддерживаются основные типы,,,, `string` `number` `boolean` `void` и `undefined` .</span><span class="sxs-lookup"><span data-stu-id="a9f81-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="a9f81-155">Типы объединения, используемые в качестве возвращаемых типов, действуют с теми же ограничениями, что и при использовании в качестве параметров сценария.</span><span class="sxs-lookup"><span data-stu-id="a9f81-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="a9f81-156">Типы массивов разрешены, если они имеют тип `string` , `number` или `boolean` .</span><span class="sxs-lookup"><span data-stu-id="a9f81-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="a9f81-157">Они также разрешены, если тип является поддерживаемым объединением или поддерживаемым типом литерала.</span><span class="sxs-lookup"><span data-stu-id="a9f81-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="a9f81-158">Типы объектов, используемые в качестве возвращаемых типов, действуют с теми же ограничениями, что и при использовании в качестве параметров сценария.</span><span class="sxs-lookup"><span data-stu-id="a9f81-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="a9f81-159">Неявная типизация поддерживается, несмотря на то, что они должны следовать тем же правилам, что и определенный тип.</span><span class="sxs-lookup"><span data-stu-id="a9f81-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="a9f81-160">Избегайте использования относительных ссылок</span><span class="sxs-lookup"><span data-stu-id="a9f81-160">Avoid using relative references</span></span>

<span data-ttu-id="a9f81-161">Power автоматизирует выполнение вашего сценария в выбранной книге Excel от вашего имени.</span><span class="sxs-lookup"><span data-stu-id="a9f81-161">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="a9f81-162">В этом случае книга может быть закрыта.</span><span class="sxs-lookup"><span data-stu-id="a9f81-162">The workbook might be closed when this happens.</span></span> <span data-ttu-id="a9f81-163">Любой API, зависящий от текущего состояния пользователя (например `Workbook.getActiveWorksheet` ,), не будет работать при использовании автоматизации Powering.</span><span class="sxs-lookup"><span data-stu-id="a9f81-163">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="a9f81-164">При проектировании скриптов обязательно используйте абсолютные ссылки на листы и диапазоны.</span><span class="sxs-lookup"><span data-stu-id="a9f81-164">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="a9f81-165">Приведенные ниже методы вызовут ошибку и завершатся ошибкой при вызове из скрипта в блоке автоматизации Power.</span><span class="sxs-lookup"><span data-stu-id="a9f81-165">The following methods will throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="a9f81-166">Класс</span><span class="sxs-lookup"><span data-stu-id="a9f81-166">Class</span></span> | <span data-ttu-id="a9f81-167">Метод</span><span class="sxs-lookup"><span data-stu-id="a9f81-167">Method</span></span> |
|--|--|
| [<span data-ttu-id="a9f81-168">Chart</span><span class="sxs-lookup"><span data-stu-id="a9f81-168">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="a9f81-169">Range</span><span class="sxs-lookup"><span data-stu-id="a9f81-169">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="a9f81-170">Workbook</span><span class="sxs-lookup"><span data-stu-id="a9f81-170">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="a9f81-171">Workbook</span><span class="sxs-lookup"><span data-stu-id="a9f81-171">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="a9f81-172">Workbook</span><span class="sxs-lookup"><span data-stu-id="a9f81-172">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="a9f81-173">Workbook</span><span class="sxs-lookup"><span data-stu-id="a9f81-173">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` |
| [<span data-ttu-id="a9f81-174">Workbook</span><span class="sxs-lookup"><span data-stu-id="a9f81-174">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="a9f81-175">Workbook</span><span class="sxs-lookup"><span data-stu-id="a9f81-175">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |
| [<span data-ttu-id="a9f81-176">Worksheet</span><span class="sxs-lookup"><span data-stu-id="a9f81-176">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `activate` |

## <a name="example"></a><span data-ttu-id="a9f81-177">Пример</span><span class="sxs-lookup"><span data-stu-id="a9f81-177">Example</span></span>

<span data-ttu-id="a9f81-178">На следующем снимке экрана показан процесс автоматизации Power, который срабатывает при назначении вопроса [GitHub](https://github.com/) .</span><span class="sxs-lookup"><span data-stu-id="a9f81-178">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="a9f81-179">Поток выполняет сценарий, который добавляет ошибку в таблицу в книге Excel.</span><span class="sxs-lookup"><span data-stu-id="a9f81-179">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="a9f81-180">Если в этой таблице имеется пять или более проблем, посылается напоминание по электронной почте.</span><span class="sxs-lookup"><span data-stu-id="a9f81-180">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![Пример процесса, показанный в редакторе автоматизации управления питанием.](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="a9f81-182">`main`Функция скрипта ЗАДАЕТ идентификатор вопроса и заголовок вопроса в качестве входных параметров, а скрипт возвращает количество строк в таблице "ошибка".</span><span class="sxs-lookup"><span data-stu-id="a9f81-182">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="a9f81-183">См. также</span><span class="sxs-lookup"><span data-stu-id="a9f81-183">See also</span></span>

- [<span data-ttu-id="a9f81-184">Запуск сценариев Office в Excel в Интернете с помощью Power автоматизиру</span><span class="sxs-lookup"><span data-stu-id="a9f81-184">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="a9f81-185">Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate</span><span class="sxs-lookup"><span data-stu-id="a9f81-185">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="a9f81-186">Основные сведения о сценариях Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="a9f81-186">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="a9f81-187">Начало работы с Power Automate</span><span class="sxs-lookup"><span data-stu-id="a9f81-187">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="a9f81-188">Справочная документация по Microsoft Online Connector (бизнес)</span><span class="sxs-lookup"><span data-stu-id="a9f81-188">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
