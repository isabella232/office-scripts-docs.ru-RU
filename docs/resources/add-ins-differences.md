---
title: Различия между сценариями Office и надстройками Office
description: Различия в поведении и API между сценариями Office и надстройками Office.
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 2290d4e34b7a7286d67443de9e9c64bad4fcd4b7
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978728"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="62d6f-103">Различия между сценариями Office и надстройками Office</span><span class="sxs-lookup"><span data-stu-id="62d6f-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="62d6f-104">Надстройки Office и сценарии Office широко распространены.</span><span class="sxs-lookup"><span data-stu-id="62d6f-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="62d6f-105">Они оба предоставляют автоматизированный контроль над книгой Excel с `Excel` помощью пространства имен API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="62d6f-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="62d6f-106">Однако в их области более ограничены скрипты Office.</span><span class="sxs-lookup"><span data-stu-id="62d6f-106">However, Office Scripts are more limited in their scope.</span></span>

![Схема из четырех квадрантов, в которой показаны области фокуса для различных решений по расширению Office.](../images/office-programmability-diagram.png)

<span data-ttu-id="62d6f-109">Сценарии Office выполняются с помощью нажатия кнопки вручную или в качестве этапа [автоматизированного управления](https://flow.microsoft.com/), в то время как надстройки Office остаются открытыми при открытии их областей задач.</span><span class="sxs-lookup"><span data-stu-id="62d6f-109">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="62d6f-110">Это означает, что надстройки могут сохранять состояние во время сеанса, в то время как сценарии Office не поддерживают внутреннее состояние между запусками.</span><span class="sxs-lookup"><span data-stu-id="62d6f-110">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="62d6f-111">Если расширение Excel должно превышать возможности платформы сценариев, посетите [документацию по надстройкам Office](/office/dev/add-ins) , чтобы узнать больше о надстройках Office.</span><span class="sxs-lookup"><span data-stu-id="62d6f-111">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="62d6f-112">В оставшейся части этой статьи описываются основные различия между надстройками Office и сценариями Office.</span><span class="sxs-lookup"><span data-stu-id="62d6f-112">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="62d6f-113">Поддержка платформы</span><span class="sxs-lookup"><span data-stu-id="62d6f-113">Platform Support</span></span>

<span data-ttu-id="62d6f-114">Надстройки Office на нескольких платформах.</span><span class="sxs-lookup"><span data-stu-id="62d6f-114">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="62d6f-115">Они работают на компьютерах с Windows, Mac, iOS и на веб-платформах и предоставляют одни и те же возможности.</span><span class="sxs-lookup"><span data-stu-id="62d6f-115">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="62d6f-116">Любое исключение из этого параметра отмечается в документации по отдельному API.</span><span class="sxs-lookup"><span data-stu-id="62d6f-116">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="62d6f-117">В настоящее время скрипты Office поддерживаются только для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="62d6f-117">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="62d6f-118">Все операции записи, редактирования и запуска выполняются на веб-платформе.</span><span class="sxs-lookup"><span data-stu-id="62d6f-118">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="62d6f-119">Интерфейсы API</span><span class="sxs-lookup"><span data-stu-id="62d6f-119">APIs</span></span>

<span data-ttu-id="62d6f-120">Сценарии Office поддерживают большинство API JavaScript для Excel, что означает, что между этими платформами существует множество функций.</span><span class="sxs-lookup"><span data-stu-id="62d6f-120">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="62d6f-121">Существует два исключения: события и общие API.</span><span class="sxs-lookup"><span data-stu-id="62d6f-121">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="62d6f-122">События</span><span class="sxs-lookup"><span data-stu-id="62d6f-122">Events</span></span>

<span data-ttu-id="62d6f-123">Сценарии Office не поддерживают [события](/office/dev/add-ins/excel/excel-add-ins-events).</span><span class="sxs-lookup"><span data-stu-id="62d6f-123">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="62d6f-124">Каждый сценарий выполняет код в отдельном `main` методе, а затем завершается.</span><span class="sxs-lookup"><span data-stu-id="62d6f-124">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="62d6f-125">Он не активируется повторно при срабатывании событий, поэтому не может регистрировать события.</span><span class="sxs-lookup"><span data-stu-id="62d6f-125">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="62d6f-126">Общие API</span><span class="sxs-lookup"><span data-stu-id="62d6f-126">Common APIs</span></span>

<span data-ttu-id="62d6f-127">Скрипты Office не могут использовать [Общие API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="62d6f-127">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="62d6f-128">Если требуется проверка подлинности, диалоговые окна или другие функции, которые поддерживаются только общими API, скорее всего, потребуется создать надстройку Office, а не сценарий Office.</span><span class="sxs-lookup"><span data-stu-id="62d6f-128">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="62d6f-129">См. также</span><span class="sxs-lookup"><span data-stu-id="62d6f-129">See also</span></span>

- [<span data-ttu-id="62d6f-130">Сценарии Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="62d6f-130">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="62d6f-131">Различия между сценариями Office и макросами VBA</span><span class="sxs-lookup"><span data-stu-id="62d6f-131">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="62d6f-132">Устранение неполадок в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="62d6f-132">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="62d6f-133">Создание надстройки области задач Excel</span><span class="sxs-lookup"><span data-stu-id="62d6f-133">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
