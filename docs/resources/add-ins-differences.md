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
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Различия между сценариями Office и надстройками Office

Надстройки Office и сценарии Office широко распространены. Они оба предоставляют автоматизированный контроль над книгой Excel с `Excel` помощью пространства имен API JavaScript для Office. Однако в их области более ограничены скрипты Office.

Сценарии Office выполняются с нажатием кнопки вручную, а надстройки Office основываются на взаимодействии с пользователем и остаются во время использования книги. Если расширение Excel должно превышать возможности платформы сценариев, посетите [документацию по надстройкам Office](/office/dev/add-ins) , чтобы узнать больше о надстройках Office.

В оставшейся части этой статьи описываются основные различия между надстройками Office и сценариями Office.

## <a name="platform-support"></a>Поддержка платформы

Надстройки Office на нескольких платформах. Они работают на компьютерах с Windows, Mac, iOS и на веб-платформах и предоставляют одни и те же возможности. Любое исключение из этого параметра отмечается в документации по отдельному API.

В настоящее время скрипты Office поддерживаются только для Excel в Интернете. Все операции записи, редактирования и запуска выполняются на веб-платформе.

## <a name="apis"></a>Интерфейсы API

Сценарии Office поддерживают большинство API JavaScript для Excel, что означает, что между этими платформами существует множество функций. Существует два исключения: события и общие API.

### <a name="events"></a>События

Сценарии Office не поддерживают [события](/office/dev/add-ins/excel/excel-add-ins-events). Каждый сценарий выполняет код в отдельном `main` методе, а затем завершается. Он не активируется повторно при срабатывании событий, поэтому не может регистрировать события.

### <a name="common-apis"></a>Общие API

Скрипты Office не могут использовать [Общие API](/javascript/api/office). Если требуется проверка подлинности, диалоговые окна или другие функции, которые поддерживаются только общими API, скорее всего, потребуется создать надстройку Office, а не сценарий Office.

## <a name="see-also"></a>См. также

- [Сценарии Office в Excel в Интернете](../overview/excel.md)
- [Устранение неполадок сценариев Office](../testing/troubleshooting.md)
- [Создание надстройки области задач Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)