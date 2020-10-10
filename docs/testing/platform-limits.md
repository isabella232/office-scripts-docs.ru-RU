---
title: Требования к платформе и требования для сценариев Office
description: Пределы ресурсов и поддержка браузеров для сценариев Office при использовании с Excel в Интернете
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: df468192f443b912e26411e46c9f953e046e55ec
ms.sourcegitcommit: 42fa3b629c93930b4e73e9c4c01d0c8bdf6d7487
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/09/2020
ms.locfileid: "48411559"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="3bd24-103">Требования к платформе и требования для сценариев Office</span><span class="sxs-lookup"><span data-stu-id="3bd24-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="3bd24-104">Существуют некоторые ограничения платформы, на которые следует обратить внимание при разработке сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="3bd24-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="3bd24-105">В этой статье приведены сведения о поддержке браузеров и данных для скриптов Office для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="3bd24-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="3bd24-106">Поддержка браузеров</span><span class="sxs-lookup"><span data-stu-id="3bd24-106">Browser support</span></span>

<span data-ttu-id="3bd24-107">Сценарии Office работают в любом браузере, [поддерживающем Office для Интернета](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="3bd24-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="3bd24-108">Однако некоторые функции JavaScript не поддерживаются в Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="3bd24-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="3bd24-109">Все функции, реализованные в [ES6 или более поздней версии](https://www.w3schools.com/Js/js_es6.asp) , не будут работать с IE 11.</span><span class="sxs-lookup"><span data-stu-id="3bd24-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="3bd24-110">Если пользователи в вашей организации по-прежнему используют этот браузер, обязательно протестируйте сценарии в этой среде при их совместном использовании.</span><span class="sxs-lookup"><span data-stu-id="3bd24-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="3bd24-111">Сторонние файлы cookie</span><span class="sxs-lookup"><span data-stu-id="3bd24-111">Third-party cookies</span></span>

<span data-ttu-id="3bd24-112">Для отображения вкладки " **Автоматизация** " в Excel в Интернете необходимо, чтобы в браузере были включены сторонние файлы cookie.</span><span class="sxs-lookup"><span data-stu-id="3bd24-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="3bd24-113">Проверьте параметры браузера, если вкладка не отображается.</span><span class="sxs-lookup"><span data-stu-id="3bd24-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="3bd24-114">Если вы используете частный сеанс браузера, возможно, потребуется повторно включить этот параметр каждый раз.</span><span class="sxs-lookup"><span data-stu-id="3bd24-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="3bd24-115">Некоторые браузеры ссылаются на этот параметр как "все файлы cookie", а не как "сторонние файлы cookie".</span><span class="sxs-lookup"><span data-stu-id="3bd24-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

## <a name="data-limits"></a><span data-ttu-id="3bd24-116">Ограничения данных</span><span class="sxs-lookup"><span data-stu-id="3bd24-116">Data limits</span></span>

<span data-ttu-id="3bd24-117">Существует ряд условий, определяющих, сколько данных Excel можно переносить одновременно и сколько можно выполнить отдельные транзакции автоматизации Power.</span><span class="sxs-lookup"><span data-stu-id="3bd24-117">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="3bd24-118">Excel</span><span class="sxs-lookup"><span data-stu-id="3bd24-118">Excel</span></span>

<span data-ttu-id="3bd24-119">При совершении вызовов книги с помощью сценария в Excel для Интернета действуют следующие ограничения:</span><span class="sxs-lookup"><span data-stu-id="3bd24-119">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="3bd24-120">Количество запросов и ответов не может превышать **5 МБ**.</span><span class="sxs-lookup"><span data-stu-id="3bd24-120">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="3bd24-121">Диапазон ограничен **5 000 000 ячейками**.</span><span class="sxs-lookup"><span data-stu-id="3bd24-121">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="3bd24-122">Если при работе с большими наборами данных возникают ошибки, попробуйте использовать несколько меньших диапазонов, а не больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="3bd24-122">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="3bd24-123">Кроме того, можно использовать API, такие как [Range. жетспеЦиалцеллс](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) , для назначения определенных ячеек, а не больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="3bd24-123">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="3bd24-124">Power Automate</span><span class="sxs-lookup"><span data-stu-id="3bd24-124">Power Automate</span></span>

<span data-ttu-id="3bd24-125">При использовании сценариев Office с автоматизацией энергосбережения вы можете **200 вызовов в день**.</span><span class="sxs-lookup"><span data-stu-id="3bd24-125">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="3bd24-126">Этот лимит сбрасывается в 12:00 AM UTC.</span><span class="sxs-lookup"><span data-stu-id="3bd24-126">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="3bd24-127">В платформе автоматизации управления питанием также действуют ограничения на использование, которые можно найти в разделе [ограничения и настройка в Power автоматизирует](/power-automate/limits-and-config).</span><span class="sxs-lookup"><span data-stu-id="3bd24-127">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="3bd24-128">См. также</span><span class="sxs-lookup"><span data-stu-id="3bd24-128">See also</span></span>

- [<span data-ttu-id="3bd24-129">Устранение неполадок в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="3bd24-129">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="3bd24-130">Отмена эффектов сценария Office</span><span class="sxs-lookup"><span data-stu-id="3bd24-130">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="3bd24-131">Повышение производительности сценариев Office</span><span class="sxs-lookup"><span data-stu-id="3bd24-131">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="3bd24-132">Основные сведения о сценариях для сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="3bd24-132">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
