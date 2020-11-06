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
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Требования к платформе и требования для сценариев Office

Существуют некоторые ограничения платформы, на которые следует обратить внимание при разработке сценариев Office. В этой статье приведены сведения о поддержке браузеров и данных для скриптов Office для Excel в Интернете.

## <a name="browser-support"></a>Поддержка браузеров

Сценарии Office работают в любом браузере, [поддерживающем Office для Интернета](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Однако некоторые функции JavaScript не поддерживаются в Internet Explorer 11 (IE 11). Все функции, реализованные в [ES6 или более поздней версии](https://www.w3schools.com/Js/js_es6.asp) , не будут работать с IE 11. Если пользователи в вашей организации по-прежнему используют этот браузер, обязательно протестируйте сценарии в этой среде при их совместном использовании.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Сторонние файлы cookie

Для отображения вкладки " **Автоматизация** " в Excel в Интернете необходимо, чтобы в браузере были включены сторонние файлы cookie. Проверьте параметры браузера, если вкладка не отображается. Если вы используете частный сеанс браузера, возможно, потребуется повторно включить этот параметр каждый раз.

> [!NOTE]
> Некоторые браузеры ссылаются на этот параметр как "все файлы cookie", а не как "сторонние файлы cookie".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Инструкции по настройке параметров файлов cookie в популярных браузерах

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Ограничения данных

Существует ряд условий, определяющих, сколько данных Excel можно переносить одновременно и сколько можно выполнить отдельные транзакции автоматизации Power.

### <a name="excel"></a>Excel

При совершении вызовов книги с помощью сценария в Excel для Интернета действуют следующие ограничения:

- Количество запросов и ответов не может превышать **5 МБ**.
- Диапазон ограничен **5 000 000 ячейками**.

Если при работе с большими наборами данных возникают ошибки, попробуйте использовать несколько меньших диапазонов, а не больших диапазонов. Кроме того, можно использовать API, такие как [Range. жетспеЦиалцеллс](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) , для назначения определенных ячеек, а не больших диапазонов.

### <a name="power-automate"></a>Power Automate

При использовании сценариев Office с автоматизацией энергосбережения вы можете **200 вызовов в день**. Этот лимит сбрасывается в 12:00 AM UTC.

В платформе автоматизации управления питанием также действуют ограничения на использование, которые можно найти в разделе [ограничения и настройка в Power автоматизирует](/power-automate/limits-and-config).

## <a name="see-also"></a>См. также

- [Устранение неполадок в сценариях Office](troubleshooting.md)
- [Отмена эффектов сценария Office](undo.md)
- [Повышение производительности сценариев Office](../develop/web-client-performance.md)
- [Основные сведения о сценариях для сценариев Office в Excel в Интернете](../develop/scripting-fundamentals.md)
