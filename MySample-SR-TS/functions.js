(function(){"use strict";function e(){return Math.floor(100*Math.random())}CustomFunctions.associate("ADD",(function(e,t){return e+t})),CustomFunctions.associate("LOG",(function(e){return console.log(e),e})),CustomFunctions.associate("ECHO",(function(e){return null==e?"null":e})),CustomFunctions.associate("CREATEFORMATTEDNUMBER",(function(e,t){return{type:"FormattedNumber",basicValue:e,numberFormat:t}})),CustomFunctions.associate("CREATEPWMFORNUMBER",(function(e){return{type:Excel.CellValueType.double,basicValue:e,basicType:Excel.RangeValueType.double,properties:{Name:{type:Excel.CellValueType.string,basicValue:"Metadata for the number"}},layouts:{compact:{icon:Excel.EntityCompactLayoutIcons.airplane}}}})),CustomFunctions.associate("PLUSONEFORANY",(function(e){return"number"==typeof e?e+1:"object"==typeof e?((e.type===Excel.CellValueType.double||e.type===Excel.CellValueType.formattedNumber)&&(e.basicValue=e.basicValue+1),e):e})),CustomFunctions.associate("PLUSONEFORNUMBER",(function(e){return e+1})),CustomFunctions.associate("PLUSONEFORDOUBLECELLVALUE",(function(e){return e.basicValue=e.basicValue+1,e})),CustomFunctions.associate("PLUSONEFORFORMATTEDNUMBERCELLVALUE",(function(e){return e.basicValue=e.basicValue+1,e})),CustomFunctions.associate("TESTSTREAMING",(function(e,t,r){let o=0,a={type:"Entity",text:"Entity "+o,properties:{propNumber:{type:"Double",basicValue:123}}};const l=setInterval((()=>{o+=1,a.text="Entity "+o,r.setResult(a)}),1e3*t);r.onCanceled=()=>{clearInterval(l)}})),CustomFunctions.associate("TESTREPEATINGPARAMETER",(async function(e,t){let r="";const o=new Excel.RequestContext;let a=o.workbook.worksheets.getActiveWorksheet();for(let l=0;l<e.length;l++){const n=e[l];if(0===n&&null!=t.parameterAddresses[l]){let e=a.getRange(t.parameterAddresses[l]).load("text");await o.sync(),""==e.text[0][0]?r+="[]":r+=e.text[0][0]}else r+=n}return r})),CustomFunctions.associate("TESTCALLWRITEAPI",(async function(){return Excel.run((async e=>{e.workbook.worksheets.getActiveWorksheet().getRange("A1").values=[["Hello"]],await e.sync()})),"Write API called"})),CustomFunctions.associate("RETURNAFTERASYNCLATENCY",(function(e,t){let r=1e3*(2*Math.random()-1)+e;return new Promise((e=>{setTimeout((()=>{e(Math.floor(r))}),r)}))})),CustomFunctions.associate("RETURNAFTERSLEEP",(function(e,t){let r=(new Date).getTime(),o=null;do{o=(new Date).getTime()}while(o-r<e);return e})),CustomFunctions.associate("GETSIMPLEENTITY",(function(){console.log("Start getSimpleEntity");let e=Math.floor(100*Math.random());return{type:Excel.CellValueType.entity,text:"Random Entity "+e,properties:{randomNumber:{type:Excel.CellValueType.double,basicValue:e}}}})),CustomFunctions.associate("GETRANDOMENTITYAFTERASYNCLATENTCY",(function(e,t){console.log("Start getSimpleEntityAfterAsyncLatentcy");let r=Math.floor(100*Math.random());const o={type:Excel.CellValueType.entity,text:"Random Entity "+r,properties:{randomNumber:{type:Excel.CellValueType.double,basicValue:r}}};return new Promise((t=>{setTimeout((()=>{t(o)}),e)}))})),CustomFunctions.associate("GETRICHERROR",(function(e){console.log("Start getRichError");let t=Excel.ErrorCellValueType.value,r=null;switch(e.toLowerCase()){case"blocked":t=Excel.ErrorCellValueType.blocked,r=Excel.BlockedErrorCellValueSubType.dataTypeUnsupportedApp;break;case"busy":t=Excel.ErrorCellValueType.busy,r=Excel.BusyErrorCellValueSubType.loadingImage;break;case"calc":t=Excel.ErrorCellValueType.calc,r=Excel.CalcErrorCellValueSubType.tooDeeplyNested;break;case"connect":t=Excel.ErrorCellValueType.connect,r=Excel.ConnectErrorCellValueSubType.externalLinksAccessFailed;break;case"div0":t=Excel.ErrorCellValueType.div0;break;case"external":t=Excel.ErrorCellValueType.external,r=Excel.ExternalErrorCellValueSubType.unknown;break;case"field":t=Excel.ErrorCellValueType.field,r=Excel.FieldErrorCellValueSubType.webImageMissingFilePart;break;case"gettingdata":t=Excel.ErrorCellValueType.gettingData;break;case"notavailable":t=Excel.ErrorCellValueType.notAvailable;break;case"name":default:t=Excel.ErrorCellValueType.name;break;case"null":t=Excel.ErrorCellValueType.null;break;case"num":t=Excel.ErrorCellValueType.num,r=Excel.NumErrorCellValueSubType.arrayTooLarge;break;case"ref":t=Excel.ErrorCellValueType.ref,r=Excel.RefErrorCellValueSubType.externalLinksCalculatedRef;break;case"spill":t=Excel.ErrorCellValueType.spill,r=Excel.SpillErrorCellValueSubType.collision;break;case"timeout":t=Excel.ErrorCellValueType.timeout,r=Excel.TimeoutErrorCellValueSubType.pythonTimeoutLimitReached;break;case"value":t=Excel.ErrorCellValueType.value,r=Excel.ValueErrorCellValueSubType.coerceStringToNumberInvalid}let o={};return o=r?{type:Excel.CellValueType.error,basicType:Excel.RangeValueType.error,errorType:t,errorSubType:r}:{type:Excel.CellValueType.error,basicType:Excel.RangeValueType.error,errorType:t},o})),CustomFunctions.associate("GETCFERROR",(function(e,t){console.log("Start getCFError");let r=CustomFunctions.ErrorCode.notAvailable;switch(e.toLowerCase()){case"divisionbyzero":r=CustomFunctions.ErrorCode.divisionByZero;break;case"invalidvalue":r=CustomFunctions.ErrorCode.invalidValue;break;case"notavailable":r=CustomFunctions.ErrorCode.notAvailable}if(t)return new CustomFunctions.Error(r);{let e="Customized CF error message";return new CustomFunctions.Error(r,e)}})),CustomFunctions.associate("GETCFERRORMESSAGE",(function(e){return e.type==CustomFunctions.Error?e.message:e.type==Excel.CellValueType.error?"Not CF error but Excel error":"Not a CF error"})),CustomFunctions.associate("TESTFORMATTEDNUMBERSTREAMING",(function(t){const r={basicValue:e(),numberFormat:"0.0",type:Excel.CellValueType.formattedNumber};t.setResult(r),setInterval((async()=>{const r={basicValue:e(),numberFormat:"0.0",type:Excel.CellValueType.formattedNumber};t.setResult(r)}),2e3)}))})(),function(){const e="products",t=268436224,r="en-US";CustomFunctions.associate("GETRANDOMLINKEDENTITY",(async function(){console.log("Start getRandomLinkedEntity ...");let o=Math.floor(100*Math.random());return{type:"LinkedEntity",text:"Linked Entity "+o,id:{entityId:o.toString(),domainId:e,serviceId:t,culture:r}}})),CustomFunctions.associate("GETLINKEDENTITYBYID",(async function(o){return console.log(`Start getProductById: Fetching product with id ${o} ...`),{type:"LinkedEntity",text:"Chai",id:{entityId:o,domainId:e,serviceId:t,culture:r}}})),CustomFunctions.associate("PRODUCTLINKEDENTITYSERVICE",(async function(o){console.log(`Start productLinkedEntityService: Fetching linked entity with id ${o} ...`);const a=new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);console.log(`Fetching linked entity with id ${o} ...`);try{return function(o){const a={type:"LinkedEntity",text:"Linked Entity "+o,id:{entityId:o,domainId:e,serviceId:t,culture:r},properties:{"Product ID":{type:"String",basicValue:o.toString()},"Product Name":{type:"String",basicValue:"Chai"},"Quantity Per Unit":{type:"String",basicValue:"10 boxes x 20 bags"},"Unit Price":{type:"FormattedNumber",basicValue:18,numberFormat:"$* #,##0.00"},Discontinued:{type:"Boolean",basicValue:!1}},cardLayout:{title:{property:"Product Name"},sections:[{layout:"List",properties:["Product ID"]},{layout:"List",title:"Quantity and price",collapsible:!0,collapsed:!1,properties:["Quantity Per Unit","Unit Price"]},{layout:"List",title:"Additional information",collapsed:!0,properties:["Discontinued"]}]}};return a.properties.Image={type:"WebImage",address:"https://upload.wikimedia.org/wikipedia/commons/thumb/0/04/Masala_Chai.JPG/320px-Masala_Chai.JPG"},a.cardLayout.mainImage={property:"Image"},a.properties.Category={type:"LinkedEntity",text:"Beverages",id:{entityId:"C1",domainId:"categories",serviceId:t,culture:r}},a.cardLayout.sections[0].properties.push("Category"),a.properties.Supplier={type:"LinkedEntity",text:"Exotic Liquids",id:{entityId:"S1",domainId:"suppliers",serviceId:t,culture:r}},a.cardLayout.sections[2].properties.push("Supplier"),a}(JSON.parse(o).entityId)}catch(e){throw console.error(e),a}}))}();
//# sourceMappingURL=functions.js.map