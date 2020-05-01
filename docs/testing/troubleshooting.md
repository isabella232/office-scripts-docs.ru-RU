---
title: Устранение неполадок в сценариях Office
description: Советы и методы отладки сценариев Office, а также справочные ресурсы.
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700357"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="6a304-103">Устранение неполадок в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="6a304-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="6a304-104">Когда вы разрабатываете сценарии Office, вы можете делать ошибки.</span><span class="sxs-lookup"><span data-stu-id="6a304-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="6a304-105">Всё в порядке.</span><span class="sxs-lookup"><span data-stu-id="6a304-105">It's okay.</span></span> <span data-ttu-id="6a304-106">У нас есть инструменты, которые помогают находить проблемы и работать с ними в идеальном состоянии.</span><span class="sxs-lookup"><span data-stu-id="6a304-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="6a304-107">Журналы консоли</span><span class="sxs-lookup"><span data-stu-id="6a304-107">Console logs</span></span>

<span data-ttu-id="6a304-108">Иногда при устранении неполадок необходимо напечатать сообщения на экране.</span><span class="sxs-lookup"><span data-stu-id="6a304-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="6a304-109">Здесь можно отобразить текущее значение переменных или триггеров, которые вызываются.</span><span class="sxs-lookup"><span data-stu-id="6a304-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="6a304-110">Для этого зарегистрируете текст в консоли.</span><span class="sxs-lookup"><span data-stu-id="6a304-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> <span data-ttu-id="6a304-111">Не забудьте `load` получить данные листа `sync` и книгу перед занесением в него свойств объекта.</span><span class="sxs-lookup"><span data-stu-id="6a304-111">Don't forget to `load` worksheet data and `sync` with the workbook before logging object properties.</span></span>

<span data-ttu-id="6a304-112">Строки, в`console.log` которые передаются данные, отображаются в консоли ведения журнала редактора кода.</span><span class="sxs-lookup"><span data-stu-id="6a304-112">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="6a304-113">Чтобы включить консоль, нажмите кнопку **с многоточием** и выберите пункт **журналы...**</span><span class="sxs-lookup"><span data-stu-id="6a304-113">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="6a304-114">Журналы не влияют на книгу.</span><span class="sxs-lookup"><span data-stu-id="6a304-114">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="6a304-115">Сообщения об ошибках</span><span class="sxs-lookup"><span data-stu-id="6a304-115">Error messages</span></span>

<span data-ttu-id="6a304-116">Если при выполнении сценария Excel возникла проблема, она выдает ошибку.</span><span class="sxs-lookup"><span data-stu-id="6a304-116">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="6a304-117">Появится всплывающее сообщение с вопросом, следует ли **просматривать журналы**.</span><span class="sxs-lookup"><span data-stu-id="6a304-117">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="6a304-118">Нажмите эту кнопку, чтобы открыть консоль и отобразить все ошибки.</span><span class="sxs-lookup"><span data-stu-id="6a304-118">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="6a304-119">Справочные материалы</span><span class="sxs-lookup"><span data-stu-id="6a304-119">Help resources</span></span>

<span data-ttu-id="6a304-120">[Переполнение стека](https://stackoverflow.com/questions/tagged/office-scripts) — это сообщество разработчиков, которые могут помочь при возникновении проблем с написанием кода.</span><span class="sxs-lookup"><span data-stu-id="6a304-120">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="6a304-121">Часто вы сможете найти решение проблемы с помощью быстрого поиска переполнения стека.</span><span class="sxs-lookup"><span data-stu-id="6a304-121">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="6a304-122">Если это не так, задайте свой вопрос и пометьте его тегом "Office – Scripts".</span><span class="sxs-lookup"><span data-stu-id="6a304-122">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="6a304-123">Запомните, что вы создаете *сценарий*Office, а не *надстройку*Office.</span><span class="sxs-lookup"><span data-stu-id="6a304-123">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="6a304-124">Если возникла проблема с API JavaScript для Office, создайте проблему в репозитории GitHub [OfficeDev/Office-JS](https://github.com/OfficeDev/office-js) .</span><span class="sxs-lookup"><span data-stu-id="6a304-124">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="6a304-125">Участники группы продукта будут отвечать на проблемы и предоставлять дальнейшую помощь.</span><span class="sxs-lookup"><span data-stu-id="6a304-125">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="6a304-126">Создание ошибки в репозитории **OfficeDev/Office-JS** свидетельствует о том, что в библиотеке API JavaScript для Office обнаружен изъян, который должен быть адрес группой разработки продуктов.</span><span class="sxs-lookup"><span data-stu-id="6a304-126">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="6a304-127">При возникновении проблем с регистратором действий или редактором отправьте отзыв с помощью кнопки " **справка > отзыв** " в Excel.</span><span class="sxs-lookup"><span data-stu-id="6a304-127">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="6a304-128">См. также</span><span class="sxs-lookup"><span data-stu-id="6a304-128">See also</span></span>

- [<span data-ttu-id="6a304-129">Сценарии Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="6a304-129">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="6a304-130">Основные сведения о сценариях для сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="6a304-130">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="6a304-131">Отменить эффекты сценария Office</span><span class="sxs-lookup"><span data-stu-id="6a304-131">Undo the effects of an Office Script</span></span>](undo.md)
