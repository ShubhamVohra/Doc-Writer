"use strict";
/// 
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
Object.defineProperty(exports, "__esModule", { value: true });
var core_1 = require("@angular/core");
var http_1 = require("@angular/http");
require("rxjs/add/operator/map");
var DbConnService = (function () {
    function DbConnService(http) {
        this.http = http;
    }
    DbConnService.prototype.getAgents = function () {
        //return this.http.get('https://www.kansanmedtrip.com/getData.php?module=treatment').map(res=>res.json());
    };
    DbConnService.prototype.dropdownClicked = function (option) {
        var par1 = "Hello EY Template Designer. Yes option is clicked.";
        var par2 = "Hello EY Template Designer. No option is clicked";
        Word.run(function (context) {
            var placeholder;
            var document = context.document;
            var app = context.application.context;
            var body = document.body;
            var contentControls = document.contentControls;
            var paragraphs = body.paragraphs;
            //placeholder.appearance.ti = "BoundingBox";
            context.load(paragraphs, 'text');
            return context.sync()
                .then(function () {
                if (option == "Yes") {
                    for (var i = 0; i < paragraphs.items.length; i++) {
                        body.insertText(paragraphs.items[i].text, "End");
                    }
                }
                context.load(body);
            });
        }).catch(this.errorHandler);
    };
    DbConnService.prototype.errorHandler = function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    };
    DbConnService = __decorate([
        core_1.Injectable(),
        __metadata("design:paramtypes", [http_1.Http])
    ], DbConnService);
    return DbConnService;
}());
exports.DbConnService = DbConnService;
//# sourceMappingURL=db.conn.service.js.map