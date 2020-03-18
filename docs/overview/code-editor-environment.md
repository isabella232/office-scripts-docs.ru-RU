---
title: Среда редактора кода сценариев Office
description: Сведения о необходимых компонентах и среде для сценариев Office в Excel в Интернете.
ms.date: 01/21/2020
localization_priority: Normal
ms.openlocfilehash: 06318305e4e0091ce4fd8d1cd8130c474e18aed9
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700399"
---
# <a name="office-scripts-code-editor-environment"></a>Среда редактора кода сценариев Office

Сценарии Office создаются в [TypeScript или JavaScript](#scripting-language-typescript-or-javascript) и используют сценарии Office для взаимодействия с книгой Excel с помощью [API JavaScript](#office-scripts-javascript-api) .

## <a name="scripting-language-typescript-or-javascript"></a>Язык сценариев: TypeScript или JavaScript

Скрипты Office создаются в [TypeScript](https://www.typescriptlang.org/docs/home.html) или [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). Средство записи действий создает код в TypeScript (то есть является надмножеством объекта JavaScript). В документации по сценариям Office используется TypeScript, но если вы знакомы с JavaScript, их можно использовать вместо этого.

Сценарии Office — это в основном автономные части кода. Используется только небольшая часть функциональных возможностей TypeScript. Таким образом, вы можете редактировать сценарии без необходимости изучать сложности TypeScript. Редактор кода также обрабатывает установку, компиляцию и выполнение кода, поэтому вам не придется беспокоиться о каком-либо, кроме собственно скрипта. Можно изучить язык и создать скрипты, не имеющие предыдущих знаний по программированию. Тем не менее, если вы не знакомы с программированием, рекомендуем изучить некоторые фундаментальные сведения, прежде чем приступать к работе со сценариями Office:

- Сведения об основах JavaScript. Вам понравится такие понятия, как переменные, потоки управления, функции и типы данных. [Mozilla предоставляет хорошее, всестороннее руководство по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).
- Сведения о типах в TypeScript. TypeScript строится на JavaScript, гарантируя, что во время компиляции для вызовов и назначений методов используются подходящие типы. В документации TypeScript на [интерфейсах](https://www.typescriptlang.org/docs/handbook/interfaces.html), [классах](https://www.typescriptlang.org/docs/handbook/classes.html), [определениях типов](https://www.typescriptlang.org/docs/handbook/type-inference.html)и [совместимости типов](https://www.typescriptlang.org/docs/handbook/type-compatibility.html) будет наиболее полезна.

## <a name="office-scripts-javascript-api"></a>API JavaScript для сценариев Office

Сценарии Office используют специальную версию API JavaScript для Office, которые используются надстройками [Office](/office/dev/add-ins/overview/index). Различия между этими двумя платформами описаны в статье [различия между сценариями Office и](../resources/add-ins-differences.md#apis) статьей надстройки Office. Вы можете просмотреть все API, доступные в сценарии, в [справочной документации по API для Office](/javascript/api/office-scripts/overview).

## <a name="intellisense"></a>IntelliSense

IntelliSense — это функция редактора кода, которая помогает предотвратить опечатки и синтаксические ошибки при редактировании сценария. В нем отображаются возможные имена объектов и полей при вводе, а также встроенная документация для каждого API.

Редактор кода Excel использует ту же подсистему IntelliSense, что и Visual Studio Code. Чтобы узнать больше об этой функции, перейдите по [функциям IntelliSense в Visual Studio Code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="see-also"></a>См. также

- [Справочник по API скриптов Office](/javascript/api/office-scripts/overview)
- [Устранение неполадок сценариев Office](../testing/troubleshooting.md)
- [Использование встроенных объектов JavaScript в сценариях Office](../develop/javascript-objects.md)
