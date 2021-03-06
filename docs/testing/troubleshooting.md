---
title: Устранение неполадок в сценариях Office
description: Советы и методы отладки сценариев Office, а также справочные ресурсы.
ms.date: 10/08/2020
localization_priority: Normal
ms.openlocfilehash: 9b3f4be778f3cdb4711d1e41d4d68f87ebca8152
ms.sourcegitcommit: 42fa3b629c93930b4e73e9c4c01d0c8bdf6d7487
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/09/2020
ms.locfileid: "48411550"
---
# <a name="troubleshooting-office-scripts"></a>Устранение неполадок в сценариях Office

Когда вы разрабатываете сценарии Office, вы можете делать ошибки. Всё в порядке. У нас есть инструменты, которые помогают находить проблемы и работать с ними в идеальном состоянии.

## <a name="console-logs"></a>Журналы консоли

Иногда при устранении неполадок необходимо напечатать сообщения на экране. Здесь можно отобразить текущее значение переменных или триггеров, которые вызываются. Для этого зарегистрируете текст в консоли.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Строки, в которые передаются данные `console.log` , отображаются в консоли ведения журнала редактора кода. Чтобы включить консоль, нажмите кнопку **с многоточием** и выберите пункт **журналы...**

Журналы не влияют на книгу.

## <a name="error-messages"></a>Сообщения об ошибках

Если при выполнении сценария Excel возникла проблема, она выдает ошибку. Появится всплывающее сообщение с вопросом, следует ли **просматривать журналы**. Нажмите эту кнопку, чтобы открыть консоль и отобразить все ошибки.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Автоматизация вкладок не отображается или сценарии Office недоступны

Выполните следующие действия, чтобы устранить все неполадки, связанные с вкладкой **Автоматизация** , в Excel в Интернете.

1. [Убедитесь, что лицензия microsoft 365 включает сценарии Office](../overview/excel.md#requirements).
1. [Попросите администратора включить эту функцию](/microsoft-365/admin/manage/manage-office-scripts-settings).
1. [Убедитесь, что ваш браузер поддерживается](platform-limits.md#browser-support).
1. [Убедитесь, что сторонние файлы Cookie включены](platform-limits.md#third-party-cookies).

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a>Справочные материалы

[Переполнение стека](https://stackoverflow.com/questions/tagged/office-scripts) — это сообщество разработчиков, которые могут помочь при возникновении проблем с написанием кода. Часто вы сможете найти решение проблемы с помощью быстрого поиска переполнения стека. Если это не так, задайте свой вопрос и пометьте его тегом "Office – Scripts". Запомните, что вы создаете *сценарий*Office, а не *надстройку*Office.

Если возникла проблема с API JavaScript для Office, создайте проблему в репозитории GitHub [OfficeDev/Office-JS](https://github.com/OfficeDev/office-js) . Участники группы продукта будут отвечать на проблемы и предоставлять дальнейшую помощь. Создание ошибки в репозитории **OfficeDev/Office-JS** свидетельствует о том, что в библиотеке API JavaScript для Office обнаружен изъян, который должен быть адрес группой разработки продуктов.

При возникновении проблем с регистратором действий или редактором отправьте отзыв с помощью кнопки " **справка > отзыв** " в Excel.

## <a name="see-also"></a>См. также

- [Сценарии Office в Excel в Интернете](../overview/excel.md)
- [Основные сведения о сценариях для сценариев Office в Excel в Интернете](../develop/scripting-fundamentals.md)
- [Пределы платформы с помощью сценариев Office](platform-limits.md)
- [Повышение производительности сценариев Office](../develop/web-client-performance.md)
- [Отмена эффектов сценария Office](undo.md)
