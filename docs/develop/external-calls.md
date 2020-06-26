---
title: Поддержка вызовов внешнего API в сценариях Office
description: Поддержка и рекомендации по выполнению вызовов внешнего API в скрипте Office.
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: ec8281551cbe7c500eee40ec86067e5efbfcfc31
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878820"
---
# <a name="external-api-call-support-in-office-scripts"></a>Поддержка вызовов внешнего API в сценариях Office

Платформа сценариев Office не поддерживает вызовы [внешних API](https://developer.mozilla.org/docs/Web/API). Тем не менее, эти вызовы могут выполняться в соответствии с правильными обстоятельствами. Внешние звонки можно выполнить только через клиент Excel, а не через автоматизированное управление питанием [в нормальных условиях](#external-calls-from-power-automate).

Авторы скриптов не должны ожидать согласованного поведения при использовании внешних API во время этапа предварительной версии платформы. Это обусловлено тем, как среда выполнения JavaScript управляет взаимодействии с книгой. Скрипт может завершиться до завершения вызова API (или `Promise` он полностью разрешается). Таким образом, не полагайтесь на внешние API для критически важных сценариев.

> [!CAUTION]
> Внешние вызовы могут привести к тому, что конфиденциальные данные будут представлены нежелательным конечным точкам. Администратор может установить защиту брандмауэра для таких вызовов.

## <a name="definition-files-for-external-apis"></a>Файлы определений для внешних API

Файлы определений внешних API не входят в состав сценариев Office. Использование таких API приводит к возникновению ошибок при компиляции для отсутствующих определений. API все еще выполняются (хотя только при использовании клиента Excel), как показано в следующем сценарии:

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

## <a name="external-calls-from-power-automate"></a>Внешние вызовы от автоматизации управления питанием

При запуске скрипта с помощью Power автоматизиру все вызовы внешнего API завершатся с ошибками. Это различие между выполнением скрипта с помощью клиента Excel и автоматизацией управления. Прежде чем приступать к их созданию, обязательно проверьте свои сценарии на наличие таких ссылок.

> [!WARNING]
> Сбой внешних вызовов [Microsoft Excel Online Connector](/connectors/excelonlinebusiness) в Power Автоматизация состоит в том, чтобы помочь приподнятое существующим политикам защиты от потери данных. Тем не менее скрипты, выполняемые с помощью автоматизации автоматизации, выполняются в рамках вашей организации, а не в брандмауэрах Организации. Для дополнительной защиты от злонамеренных пользователей в этой внешней среде администратор может управлять использованием сценариев Office. Администратор может отключить соединитель Excel Online в Power автоматизирует или отключить сценарии Office для Excel в Интернете с помощью [сценариев Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).

## <a name="see-also"></a>См. также

- [Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md)