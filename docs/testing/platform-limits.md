---
title: Требования к платформе и требования для сценариев Office
description: Пределы ресурсов и поддержка браузеров для сценариев Office при использовании с Excel в Интернете
ms.date: 10/23/2020
localization_priority: Normal
ms.openlocfilehash: 61f5c55be278ae056014d3b01e4176354d913f87
ms.sourcegitcommit: d3e7681e262bdccc281fcb7b3c719494202e846b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/06/2020
ms.locfileid: "48930080"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="bac36-103">Требования к платформе и требования для сценариев Office</span><span class="sxs-lookup"><span data-stu-id="bac36-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="bac36-104">Существуют некоторые ограничения платформы, на которые следует обратить внимание при разработке сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="bac36-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="bac36-105">В этой статье приведены сведения о поддержке браузеров и данных для скриптов Office для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="bac36-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="bac36-106">Поддержка браузеров</span><span class="sxs-lookup"><span data-stu-id="bac36-106">Browser support</span></span>

<span data-ttu-id="bac36-107">Сценарии Office работают в любом браузере, [поддерживающем Office для Интернета](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="bac36-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="bac36-108">Однако некоторые функции JavaScript не поддерживаются в Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="bac36-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="bac36-109">Все функции, реализованные в [ES6 или более поздней версии](https://www.w3schools.com/Js/js_es6.asp) , не будут работать с IE 11.</span><span class="sxs-lookup"><span data-stu-id="bac36-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="bac36-110">Если пользователи в вашей организации по-прежнему используют этот браузер, обязательно протестируйте сценарии в этой среде при их совместном использовании.</span><span class="sxs-lookup"><span data-stu-id="bac36-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="bac36-111">Сторонние файлы cookie</span><span class="sxs-lookup"><span data-stu-id="bac36-111">Third-party cookies</span></span>

<span data-ttu-id="bac36-112">Для отображения вкладки " **Автоматизация** " в Excel в Интернете необходимо, чтобы в браузере были включены сторонние файлы cookie.</span><span class="sxs-lookup"><span data-stu-id="bac36-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="bac36-113">Проверьте параметры браузера, если вкладка не отображается.</span><span class="sxs-lookup"><span data-stu-id="bac36-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="bac36-114">Если вы используете частный сеанс браузера, возможно, потребуется повторно включить этот параметр каждый раз.</span><span class="sxs-lookup"><span data-stu-id="bac36-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="bac36-115">Некоторые браузеры ссылаются на этот параметр как "все файлы cookie", а не как "сторонние файлы cookie".</span><span class="sxs-lookup"><span data-stu-id="bac36-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="bac36-116">Инструкции по настройке параметров файлов cookie в популярных браузерах</span><span class="sxs-lookup"><span data-stu-id="bac36-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="bac36-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="bac36-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="bac36-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="bac36-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="bac36-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="bac36-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="bac36-120">Safari</span><span class="sxs-lookup"><span data-stu-id="bac36-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="bac36-121">Ограничения данных</span><span class="sxs-lookup"><span data-stu-id="bac36-121">Data limits</span></span>

<span data-ttu-id="bac36-122">Существует ряд условий, определяющих, сколько данных Excel можно переносить одновременно и сколько можно выполнить отдельные транзакции автоматизации Power.</span><span class="sxs-lookup"><span data-stu-id="bac36-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="bac36-123">Excel</span><span class="sxs-lookup"><span data-stu-id="bac36-123">Excel</span></span>

<span data-ttu-id="bac36-124">При совершении вызовов книги с помощью сценария в Excel для Интернета действуют следующие ограничения:</span><span class="sxs-lookup"><span data-stu-id="bac36-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="bac36-125">Количество запросов и ответов не может превышать **5 МБ**.</span><span class="sxs-lookup"><span data-stu-id="bac36-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="bac36-126">Диапазон ограничен **5 000 000 ячейками**.</span><span class="sxs-lookup"><span data-stu-id="bac36-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="bac36-127">Если при работе с большими наборами данных возникают ошибки, попробуйте использовать несколько меньших диапазонов, а не больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="bac36-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="bac36-128">Кроме того, можно использовать API, такие как [Range. жетспеЦиалцеллс](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) , для назначения определенных ячеек, а не больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="bac36-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="bac36-129">Power Automate</span><span class="sxs-lookup"><span data-stu-id="bac36-129">Power Automate</span></span>

<span data-ttu-id="bac36-130">При использовании сценариев Office с автоматизацией энергосбережения вы можете **200 вызовов в день**.</span><span class="sxs-lookup"><span data-stu-id="bac36-130">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="bac36-131">Этот лимит сбрасывается в 12:00 AM UTC.</span><span class="sxs-lookup"><span data-stu-id="bac36-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="bac36-132">В платформе автоматизации управления питанием также действуют ограничения на использование, которые можно найти в разделе [ограничения и настройка в Power автоматизирует](/power-automate/limits-and-config).</span><span class="sxs-lookup"><span data-stu-id="bac36-132">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="bac36-133">См. также</span><span class="sxs-lookup"><span data-stu-id="bac36-133">See also</span></span>

- [<span data-ttu-id="bac36-134">Устранение неполадок в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="bac36-134">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="bac36-135">Отмена эффектов сценария Office</span><span class="sxs-lookup"><span data-stu-id="bac36-135">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="bac36-136">Повышение производительности сценариев Office</span><span class="sxs-lookup"><span data-stu-id="bac36-136">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="bac36-137">Основные сведения о сценариях для сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="bac36-137">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
