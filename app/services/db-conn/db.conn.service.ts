

/// 

import { Injectable } from '@angular/core';
import { Http,Headers,RequestOptions } from '@angular/http';

import 'rxjs/add/operator/map';
declare var jquery:any;
declare var $:any;

@Injectable()
export class DbConnService { 
    
    constructor(private http:Http) {}

    getAgents(){
      
        //return this.http.get('https://www.kansanmedtrip.com/getData.php?module=treatment').map(res=>res.json());
    
    }


    dropdownClicked(option:any){
        var par1 = "Hello EY Template Designer. Yes option is clicked.";
        var par2 = "Hello EY Template Designer. No option is clicked";

        Word.run(function(context){
            let placeholder:Word.ContentControl;
            
            var document = context.document;
            var app = context.application.context;
            var body = document.body;
            var contentControls = document.contentControls;
            var paragraphs = body.paragraphs;
            //placeholder.appearance.ti = "BoundingBox";
            context.load(paragraphs,'text');
            
            
            
            return context.sync()
            .then(function(){
                if(option=="Yes"){
                    for(var i =0 ;i<paragraphs.items.length;i++){
                        body.insertText(paragraphs.items[i].text,"End");
                    }
                }
                context.load(body);
                
            });

        }).catch(this.errorHandler);
    }

    errorHandler(error: any){
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    
    
}
