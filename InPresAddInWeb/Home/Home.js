/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    /*

    function goToSlideByIndex(slideIndex) {
       
        Office.context.document.goToByIdAsync(slideIndex, Office.GoToType.Index, function (asyncResult) {
            if (asyncResult.status == "failed") {
                showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                showNotification("Navigation successful");
            }
        });
    }

    

    //функция для получения информации о смене слайдов
    //Автообновение 
    window.setInterval(function () {
        //get the current slide
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (reason) {

            // null check проверка на существование слайдов?
            if (!reason || !reason.value || !reason.value.slides) {
                return;
            }

            //получаем индекс текущего слайда
            var currentSlide = reason.value.slides[0].index;
            //get stored setting for current slide
            var storedSlideIndex = Office.context.document.settings.get("CurrentSlide");
            //проверка, изменился ли слайд
            if (currentSlide != storedSlideIndex) {
            
                //update the stored setting for current slide
                // код для отправки сообщений на сервер?:
                //$.post('/remote-url', { xml: yourXMLString });
                //показываем номер слайда(для примера):
                showNotification('Index of the slide is:', '"' + currentSlide + '"');
                showNotification('Index of the previous slide is:', '"' + storedSlideIndex + '"');
            //    goToSlideByIndex(storedSlideIndex);
                Office.context.document.settings.set("CurrentSlide", currentSlide);
                Office.context.document.settings.saveAsync(function (asyncResult) { });
            }

        });

    }, 10); //количество миллисекунд для обновления


    */

    window.setInterval(function () {
        //get the current slide
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (reason) {

            // null check
            if (!reason || !reason.value || !reason.value.slides) {
                return;
            }

            //get current slides index
            var currentSlide = reason.value.slides[0].index;

            //get stored setting for current slide
            var storedSlideIndex = Office.context.document.settings.get("CurrentSlide");
            //check if current slide and stored setting are the same
            if (currentSlide != storedSlideIndex) {
                //the slide changed - do something
                //update the stored setting for current slide
                //показываем номер слайда(для примера):
                showNotification('Index of the slide is:', '"' + currentSlide + '"');
                showNotification('Index of the previous slide is:', '"' + storedSlideIndex + '"');
                Office.context.document.settings.set("CurrentSlide", currentSlide);
                Office.context.document.settings.saveAsync(function (asyncResult) { });
            }

        });

    }, 1500);


    var messageBanner;
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

   //         $('#get-data-from-selection').click(getDataFromSelection);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        if (Office.context.document.getSelectedDataAsync) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        showNotification('The selected text is:', '"' + result.value + '"');
                    } else {
                        showNotification('Error:', result.error.message);
                    }
                }
            );
        } else {
            app.showNotification('Error:', 'Reading selection data is not supported by this host application.');
        }
    }
    
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();

