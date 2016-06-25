/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

/*
Доработать: проверка наличия аддина на слайде
перемещение влево
изменить скорость перемещения (попробовать)
конец слайдов
*/


(function () {
    "use strict";

    var messageBanner;
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason)
    {
        $(document).ready(function ()
        {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            //$('#Body').click(slideHasChanged);


   //         $('#get-data-from-selection').click(getDataFromSelection);
        });
    };


    function goToSlideByIndex(slideIndex) {

        Office.context.document.goToByIdAsync(slideIndex, Office.GoToType.Index, function (asyncResult) {
            if (asyncResult.status == "failed") {
                showNotification("Action failed with error: " + asyncResult.error.message);
            }
        });
    }

    var storedSlideIndex = 0;
    var currentSlide=0;

    //функция для получения информации о смене слайдов
    window.setInterval(function slideHasChanged() {
        //get the current slide
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (reason)
        {

            // null check проверка на существование слайдов?
            if (!reason || !reason.value || !reason.value.slides) {
                return;
            }
            
            //получаем индекс текущего слайда
            currentSlide = reason.value.slides[0].index;
            
            //получаем сохраненный индекс слайда, для первого слайда значение не определено, = null
            storedSlideIndex = Office.context.document.settings.get("CurrentSlide");
            
            Office.context.document.getActiveViewAsync(function (asyncResult) ///Слишком сложное условие!!! Изменить!!
            {
                if (asyncResult.status == "failed")
                {
                        showNotification("Action failed with error: " + asyncResult.error.message);
                }

                if (asyncResult.value !== "edit") //если не находимся в режиме редактирования
                {
                    //проверка, изменился ли слайд 
                    if /*((currentSlide != storedSlideIndex) &&*/ (currentSlide>storedSlideIndex)
                    {
                        if (storedSlideIndex != null)
                        {
                            goToSlideByIndex(storedSlideIndex);
                        }
                        // код для отправки сообщений на сервер?:
                        //$.post('/remote-url', { xml: yourXMLString });
                        //показываем номер слайда(для примера):
                        // showNotification('Index of the slide is:', '"' + currentSlide + '"');
                        
                        showNotification('Index of the previous slide is:', '"' + storedSlideIndex + '"');
                        
                        Office.context.document.settings.set("CurrentSlide", currentSlide);
                        Office.context.document.settings.saveAsync(function (asyncResult) { });
                       
                    }
                    
                }
                else 
                {
                    Office.context.document.settings.set("CurrentSlide", null);
                }
            });

        });

    }, 50); //количество миллисекунд для обновления






    /*
    // Reads data from current document selection and displays a notification
    function getDataFromSelection() 
    {
        if (Office.context.document.getSelectedDataAsync) 
        {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) 
                    {
                        showNotification('The selected text is:', '"' + result.value + '"');
                    } 
                    else 
                    {
                        showNotification('Error:', result.error.message);
                    }
                }
            );
        }
        else 
        {
            app.showNotification('Error:', 'Reading selection data is not supported by this host application.');
        }
    }
    
    */

    // Функция для показа уведомлений
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();