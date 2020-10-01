---
title: Поддержка внешнего вызова API в сценариях Office
description: Поддержка и рекомендации по выполнению вызовов внешнего API в скрипте Office.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: fa77e606e2b3ab90144507660d71561b278e82e5
ms.sourcegitcommit: ce72354381561dc167ea0092efd915642a9161b3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/30/2020
ms.locfileid: "48319632"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="ba19b-103">Поддержка внешнего вызова API в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="ba19b-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="ba19b-104">Платформа сценариев Office не поддерживает вызовы [внешних API](https://developer.mozilla.org/docs/Web/API).</span><span class="sxs-lookup"><span data-stu-id="ba19b-104">The Office Scripts platform doesn't support calls to [external APIs](https://developer.mozilla.org/docs/Web/API).</span></span> <span data-ttu-id="ba19b-105">Тем не менее, эти вызовы могут выполняться в соответствии с правильными обстоятельствами.</span><span class="sxs-lookup"><span data-stu-id="ba19b-105">However, these calls can be run under the right circumstances.</span></span> <span data-ttu-id="ba19b-106">Внешние звонки можно выполнить только через клиент Excel, а не через автоматизированное управление питанием [в нормальных условиях](#external-calls-from-power-automate).</span><span class="sxs-lookup"><span data-stu-id="ba19b-106">External calls can be only be made through the Excel client, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

<span data-ttu-id="ba19b-107">Авторы скриптов не должны ожидать согласованного поведения при использовании внешних API во время этапа предварительной версии платформы.</span><span class="sxs-lookup"><span data-stu-id="ba19b-107">Script authors shouldn't expect consistent behavior when using external APIs during the platform's preview phase.</span></span> <span data-ttu-id="ba19b-108">Это обусловлено тем, как среда выполнения JavaScript управляет взаимодействии с книгой.</span><span class="sxs-lookup"><span data-stu-id="ba19b-108">This is due how the JavaScript runtime manages interacting with the workbook.</span></span> <span data-ttu-id="ba19b-109">Скрипт может завершиться до завершения вызова API (или `Promise` он полностью разрешается).</span><span class="sxs-lookup"><span data-stu-id="ba19b-109">The script may end before the API call completes (or its `Promise` is fully resolved).</span></span> <span data-ttu-id="ba19b-110">Таким образом, не полагайтесь на внешние API для критически важных сценариев.</span><span class="sxs-lookup"><span data-stu-id="ba19b-110">As such, do not rely on external APIs for critical script scenarios.</span></span>

> [!CAUTION]
> <span data-ttu-id="ba19b-111">Внешние вызовы могут привести к тому, что конфиденциальные данные будут представлены нежелательным конечным точкам.</span><span class="sxs-lookup"><span data-stu-id="ba19b-111">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="ba19b-112">Администратор может установить защиту брандмауэра для таких вызовов.</span><span class="sxs-lookup"><span data-stu-id="ba19b-112">Your admin can establish firewall protection against such calls.</span></span>

## <a name="definition-files-for-external-apis"></a><span data-ttu-id="ba19b-113">Файлы определений для внешних API</span><span class="sxs-lookup"><span data-stu-id="ba19b-113">Definition files for external APIs</span></span>

<span data-ttu-id="ba19b-114">Файлы определений внешних API не входят в состав сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="ba19b-114">The definition files for external APIs aren't included with Office Scripts.</span></span> <span data-ttu-id="ba19b-115">Использование таких API приводит к возникновению ошибок при компиляции для отсутствующих определений.</span><span class="sxs-lookup"><span data-stu-id="ba19b-115">The use of such APIs generates compile-time errors for missing definitions.</span></span> <span data-ttu-id="ba19b-116">API все еще выполняются (хотя только при использовании клиента Excel), как показано в следующем сценарии:</span><span class="sxs-lookup"><span data-stu-id="ba19b-116">The APIs still run (though only when run through the Excel client), as shown in the following script:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="ba19b-117">Внешние вызовы от автоматизации управления питанием</span><span class="sxs-lookup"><span data-stu-id="ba19b-117">External calls from Power Automate</span></span>

<span data-ttu-id="ba19b-118">При запуске скрипта с помощью Power автоматизиру все вызовы внешнего API завершатся с ошибками.</span><span class="sxs-lookup"><span data-stu-id="ba19b-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="ba19b-119">Это различие между выполнением скрипта с помощью клиента Excel и автоматизацией управления.</span><span class="sxs-lookup"><span data-stu-id="ba19b-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="ba19b-120">Прежде чем приступать к их созданию, обязательно проверьте свои сценарии на наличие таких ссылок.</span><span class="sxs-lookup"><span data-stu-id="ba19b-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="ba19b-121">Сбой внешних вызовов [Microsoft Excel Online Connector](/connectors/excelonlinebusiness) в Power Автоматизация состоит в том, чтобы помочь приподнятое существующим политикам защиты от потери данных.</span><span class="sxs-lookup"><span data-stu-id="ba19b-121">The failure of external calls [Excel Online connector](/connectors/excelonlinebusiness) in Power Automate is there to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="ba19b-122">Тем не менее скрипты, выполняемые с помощью автоматизации автоматизации, выполняются в рамках вашей организации, а не в брандмауэрах Организации.</span><span class="sxs-lookup"><span data-stu-id="ba19b-122">However, the scripts run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="ba19b-123">Для дополнительной защиты от злонамеренных пользователей в этой внешней среде администратор может управлять использованием сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="ba19b-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="ba19b-124">Администратор может отключить соединитель Excel Online в Power автоматизирует или отключить сценарии Office для Excel в Интернете с помощью [сценариев Office](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="ba19b-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="ba19b-125">См. также</span><span class="sxs-lookup"><span data-stu-id="ba19b-125">See also</span></span>

- [<span data-ttu-id="ba19b-126">Использование встроенных объектов JavaScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="ba19b-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)