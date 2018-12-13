'use strict';

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#set-color').click(setColor);
        });
    };

    function setColor() {
        Excel.run(function (context) {
            var sheet = context.workbook.worksheets.getItem("shipping_instructions");
            var range = sheet.getRange("N8");
            //range.format.fill.color = 'green';
            range.load(['address', 'values']);
            return context.sync().then(function () {
                console.log("", range.values[0][0] + " Address:" + range.address);
            });
            //research how to save the data to txt file
            
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    
})();
function hide_desc() {
    //toggle localstorage value

    if (Boolean(localStorage.desc)) {        
        localStorage.desc = "";
    }
    else {        
        localStorage.desc = "description_hidden";
    }
    set_desc_state();
}
function set_desc_state() {
    //toggle localstorage value

    if (Boolean(localStorage.desc)) {
        $(".description").addClass("hide");
        $(".desc.show").show();        
    }
    else {
        $(".desc.show").hide();
        $(".description").removeClass("hide");        
    }

}