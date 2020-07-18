---
title: Вызов сценариев из активированного вручную потока Power Automate
description: В этом руководстве рассказывается об использовании сценариев Office в Power Automate с помощью триггера с ручным срабатыванием.
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: 70fca2620973ecefe9eda40f02e28f064b713677
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160441"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a>Вызов сценариев из активированного вручную потока Power Automate (предварительный просмотр)

В этом руководстве объясняется, как запускать сценарий Office для Excel в Интернете с помощью [Power Automate](https://flow.microsoft.com).

## <a name="prerequisites"></a>Предварительные требования

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> В этом руководстве предполагается, что вы прочитали руководство [Запись, изменение и создание сценариев Office для Excel в Интернете](excel-tutorial.md).

## <a name="prepare-the-workbook"></a>Подготовка книги

В Power Automate для доступа к компонентам книги нельзя использовать такие относительные ссылки, как `Workbook.getActiveWorksheet`. Поэтому нужно использовать книгу и лист с именами, на которые может ссылаться Power Automate.

1. Создайте новую книгу под названием **MyWorkbook**.

2. В книге **MyWorkbook** создайте лист под названием **TutorialWorksheet**.

## <a name="create-an-office-script"></a>Создание сценария Office

1. Откройте вкладку **Автоматизация** и запустите **Редактор кода**.

2. Выберите **Новый сценарий**.

3. Замените сценарий по умолчанию следующим сценарием. Этот сценарий добавляет текущую дату и время в первые две ячейки листа **TutorialWorksheet**.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. Переименуйте сценарий в **Установка даты и времени**. Нажмите на имя сценария, чтобы изменить его.

5. Сохраните сценарий, нажав кнопку **Сохранить сценарий**.

## <a name="create-an-automated-workflow-with-power-automate"></a>Создание автоматизированного рабочего процесса с помощью Power Automate

1. Войдите на [сайт Power Automate](https://flow.microsoft.com).

2. В меню, которое отображается в левой части экрана, нажмите клавишу **Создать**. Откроется список способов создания новых рабочих процессов.

    ![Кнопка "Создать" в Power Automate.](../images/power-automate-tutorial-1.png)

3. В разделе **Создание нового** выберите пункт **Мгновенный поток**. В результате будет создан активированный вручную рабочий процесс.

    ![Способ мгновенного потока для создания нового рабочего процесса.](../images/power-automate-tutorial-2.png)

4. В открывшемся диалоговом окне введите имя для своего потока в поле **Имя потока**, выберите **Запустить поток вручную** из списка вариантов в разделе **Выбор способа запуска потока**и нажмите **Создать**.

    ![Способ запуска потока вручную для создания нового мгновенного потока.](../images/power-automate-tutorial-3.png)

    Обратите внимание: запускаемый вручную поток — это лишь один из многих типов потоков. В следующем руководстве описывается создание потока, который будет выполняться автоматически при получении вами сообщения электронной почты.

5. Нажмите клавишу **Следующий шаг**.

6. Откройте вкладку **Стандартные**, а затем выберите **Excel Online (бизнес)**.

    ![Функция Power Automate для Excel Online (бизнес).](../images/power-automate-tutorial-4.png)

7. В разделе **Действия** выберите **Запустить сценарий (предварительный просмотр)**.

    ![Вариант действия Power Automate "Запуск сценария" (предварительный просмотр).](../images/power-automate-tutorial-5.png)

8. Определите указанные ниже параметры для соединителя **Запуск сценария**.

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: MyWorkbook.xlsx
    - **Сценарий**: Установка даты и времени

    ![Параметры соединителя для запуска сценария в Power Automate.](../images/power-automate-tutorial-6.png)

9. Нажмите **Сохранить**.

Теперь ваш поток готов к запуску с помощью Power Automate. Вы можете проверить его с помощью кнопки **Тест** в редакторе потока или выполнить остальные действия согласно руководству, чтобы запустить поток из вашей коллекции потоков.

## <a name="run-the-script-through-power-automate"></a>Запуск сценария с помощью Power Automate

1. На главной странице Power Automate выберите **Мои потоки**.

    ![Кнопка "Мои потоки" в Power Automate.](../images/power-automate-tutorial-7.png)

2. Выберите **Мой учебный поток** из списка во вкладке **Мои потоки**. При этом будут показаны подробные сведения о потоке, который мы создали ранее.

3. Нажмите кнопку **Запуск**.

    ![Кнопка "Запуск" в Power Automate.](../images/power-automate-tutorial-8.png)

4. Появится панель задач для запуска потока. Когда будет предложено выполнить **Вход** в Excel Online, нажмите кнопку **Продолжить**.

5. Щелкните **Запустить поток**. При этом запустится поток, выполняющий связанный сценарий Office.

6. Нажмите кнопку **Готово**. Вы можете заметить, что раздел **Запуски** соответствующим образом обновлен.

7. Обновите страницу, чтобы увидеть результаты работы Power Automate. После этого перейдите в книгу, где должны отобразиться обновленные ячейки. В случае неудачи проверьте параметры этого потока и запустите его еще раз.

    ![В результатах работы Power Automate показано успешное выполнение потока.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a>Дальнейшие действия

Прочитайте раздел руководства [Передача данных сценариям в автоматически запускаемом потоке Power Automate](excel-power-automate-trigger.md). В нем рассказывается о том, как передать данные из службы рабочего процесса в ваш сценарий Office и запустить поток Power Automate при возникновении определенных событий.
