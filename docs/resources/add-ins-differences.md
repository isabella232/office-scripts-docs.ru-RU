---
title: Различия между сценариями Office и надстройками Office
description: Различия в поведении и API между сценариями Office и надстройками Office.
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700396"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="a52ee-103">Различия между сценариями Office и надстройками Office</span><span class="sxs-lookup"><span data-stu-id="a52ee-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="a52ee-104">Надстройки Office и сценарии Office широко распространены.</span><span class="sxs-lookup"><span data-stu-id="a52ee-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="a52ee-105">Они оба предоставляют автоматизированный контроль над книгой Excel с `Excel` помощью пространства имен API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="a52ee-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="a52ee-106">Однако в их области более ограничены скрипты Office.</span><span class="sxs-lookup"><span data-stu-id="a52ee-106">However, Office Scripts are more limited in their scope.</span></span>

<span data-ttu-id="a52ee-107">Сценарии Office выполняются с нажатием кнопки вручную, а надстройки Office основываются на взаимодействии с пользователем и остаются во время использования книги.</span><span class="sxs-lookup"><span data-stu-id="a52ee-107">Office Scripts run to completion with a manual button press, whereas Office Add-ins rely on user interaction and persist while the workbook is in use.</span></span> <span data-ttu-id="a52ee-108">Если расширение Excel должно превышать возможности платформы сценариев, посетите [документацию по надстройкам Office](/office/dev/add-ins) , чтобы узнать больше о надстройках Office.</span><span class="sxs-lookup"><span data-stu-id="a52ee-108">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="a52ee-109">В оставшейся части этой статьи описываются основные различия между надстройками Office и сценариями Office.</span><span class="sxs-lookup"><span data-stu-id="a52ee-109">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="a52ee-110">Поддержка платформы</span><span class="sxs-lookup"><span data-stu-id="a52ee-110">Platform Support</span></span>

<span data-ttu-id="a52ee-111">Надстройки Office на нескольких платформах.</span><span class="sxs-lookup"><span data-stu-id="a52ee-111">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="a52ee-112">Они работают на компьютерах с Windows, Mac, iOS и на веб-платформах и предоставляют одни и те же возможности.</span><span class="sxs-lookup"><span data-stu-id="a52ee-112">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="a52ee-113">Любое исключение из этого параметра отмечается в документации по отдельному API.</span><span class="sxs-lookup"><span data-stu-id="a52ee-113">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="a52ee-114">В настоящее время скрипты Office поддерживаются только для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="a52ee-114">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="a52ee-115">Все операции записи, редактирования и запуска выполняются на веб-платформе.</span><span class="sxs-lookup"><span data-stu-id="a52ee-115">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="a52ee-116">Интерфейсы API</span><span class="sxs-lookup"><span data-stu-id="a52ee-116">APIs</span></span>

<span data-ttu-id="a52ee-117">Сценарии Office поддерживают большинство API JavaScript для Excel, что означает, что между этими платформами существует множество функций.</span><span class="sxs-lookup"><span data-stu-id="a52ee-117">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="a52ee-118">Существует два исключения: события и общие API.</span><span class="sxs-lookup"><span data-stu-id="a52ee-118">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="a52ee-119">События</span><span class="sxs-lookup"><span data-stu-id="a52ee-119">Events</span></span>

<span data-ttu-id="a52ee-120">Сценарии Office не поддерживают [события](/office/dev/add-ins/excel/excel-add-ins-events).</span><span class="sxs-lookup"><span data-stu-id="a52ee-120">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="a52ee-121">Каждый сценарий выполняет код в отдельном `main` методе, а затем завершается.</span><span class="sxs-lookup"><span data-stu-id="a52ee-121">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="a52ee-122">Он не активируется повторно при срабатывании событий, поэтому не может регистрировать события.</span><span class="sxs-lookup"><span data-stu-id="a52ee-122">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="a52ee-123">Общие API</span><span class="sxs-lookup"><span data-stu-id="a52ee-123">Common APIs</span></span>

<span data-ttu-id="a52ee-124">Скрипты Office не могут использовать [Общие API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a52ee-124">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="a52ee-125">Если требуется проверка подлинности, диалоговые окна или другие функции, которые поддерживаются только общими API, скорее всего, потребуется создать надстройку Office, а не сценарий Office.</span><span class="sxs-lookup"><span data-stu-id="a52ee-125">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="a52ee-126">См. также</span><span class="sxs-lookup"><span data-stu-id="a52ee-126">See also</span></span>

- [<span data-ttu-id="a52ee-127">Сценарии Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="a52ee-127">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="a52ee-128">Устранение неполадок сценариев Office</span><span class="sxs-lookup"><span data-stu-id="a52ee-128">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="a52ee-129">Создание надстройки области задач Excel</span><span class="sxs-lookup"><span data-stu-id="a52ee-129">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)