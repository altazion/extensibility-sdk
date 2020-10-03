export module Altazion {

    export class Office {
        private static rootAppUrl : string = "";

        public static async init() : Promise<any> {

            var _rootAppUrl = Altazion.Office.getParameterByName("$from");
            if (_rootAppUrl == "") _rootAppUrl = "https://office.altatzion.com/";
            window.addEventListener("resize", function (e) {
                setTimeout(Altazion.Office.refreshSize, 150);
            });
            document.addEventListener("DOMContentLoaded", function (e) {
                setTimeout(Altazion.Office.refreshSize, 150);
            });
            window.addEventListener("load", function (e) {
                setTimeout(Altazion.Office.refreshSize, 150);
            });
            
            Altazion.Office.refreshSize();

            return await new Promise<any>(resolve => setTimeout(() => resolve( { } ), 50));
        }

        public static refreshSize() : void {
            if (parent != null) {
                var height = document.body.scrollHeight;
                if(document.documentElement!=null && document.documentElement!=undefined)
                    height = document.documentElement.scrollHeight;
                parent.postMessage("_e_sizing_;" + height, Altazion.Office.rootAppUrl);
            }
        }

        public static getParameterByName(name : string) : string {
            name = name.replace(/[\[]/, "\\\[").replace(/[\]]/, "\\\]").replace("$", "\\$");
            var regexS = "[\\?&]" + name + "=([^&#]*)";
            var regex = new RegExp(regexS);
            var results = regex.exec(window.location.search);
            if (results == null)
                return "";
            else
                return decodeURIComponent(results[1].replace(/\+/g, " "));
        }
    }

}