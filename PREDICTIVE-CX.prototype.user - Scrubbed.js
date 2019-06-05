// ==UserScript==
// @name         PREDICTIVE-CX
// @namespace    somenamespace
// @version      0.1.4beta
// @description  N/A
// @author       jesfung, fstennet
// @include      website
// @grant        GM.xmlHttpRequest
// @grant        GM_getResourceURL
// @connect      *
// ==/UserScript==

(function() {
    'use strict';

        (function() {

            function getFormDigest() {
                GM.xmlHttpRequest({
                    method: "POST",
                    url: "http://SharepointSite/_api/contextinfo",
                    headers: {
                        "accept": "application/json;odata=verbose",
                    },
                    onload: function(response) {
                        console.log('getFormDigest from SHAREPOINT: ', response.response)

                        var aaa = JSON.parse(response.response)

                        var FormDigestValue = aaa.d.GetContextWebInformation.FormDigestValue

                        var main_frame = document;
                        var body = main_frame.getElementsByTagName("body")[0];
                        var canvas = document.createElement("div");
                        canvas.setAttribute("id","mainForm");
                        canvas.setAttribute("style","visibility:hidden;position:fixed;overflow:scroll;top:100px;left:30px;background:#f5fafc;color:#1b527a;border-style:solid;border-width:1px;border-color:#1b527a;min-width:50px;min-heigth=50px;max-height:400px;z-index:1000;padding:10px;");


                        var labelq1 = ` 0 is not at all likely, 10 is extremely likely.`;
                        var labelq2 = ` a`;
                        var labelq3 = `1 is very difficult , 5 is very easy.`;

                        canvas.innerHTML = `
<form id='main_form'>
<div>
<b>PREDICTIVE CX</b><br />
<div>
<b>Service Request #</b><br/>
<input type="text" value="${document.getElementsByClassName('dynamic-caseNumberHeader')[0].innerHTML}" maxlength="9" id="ServiceRequestNumber" title="SR # (9 digits)" disabled="disabled">
</div><br/>
<div>
<b>Evaluated CSE</b><br />
<input type="text" value="${document.querySelector('[title="Owner"]').firstChild.innerHTML}" id="CaseOwnerID" title="Select the Case Owner name from SYKES Directory. Enter a name or email address..." disabled="disabled">
</div><br />
<div>
<b>01) How likely is your customer to recommend Cisco Technical Services to a friend or colleague?</b><br />
<pre>${labelq1}</pre><br />
<select id="dropdown_nps" title="${labelq1}">
<option value="" selected="selected"></option>
<option value="0">0</option>
<option value="1">1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7">7</option>
<option value="8">8</option>
<option value="9">9</option>
<option value="10">10</option>
</select><br />
</div><br />
<div>
<b>02) How do you think the customer feels about this case? </b><br/><br/>
<select id="dropdown_sentiment" title="${labelq2}">
<option value="" selected="selected"></option>
<option value="SAD">SAD</option>
<option value="NEUTRAL">NEUTRAL</option>
<option value="HAPPY">HAPPY</option>
</select><br/>
</div><br/>
<div>
<b>03) How easy has it been working with you? </b><br />
<pre>${labelq3}</pre> <br />
<select id="dropdown_effort" title="${labelq3}">
<option value="" selected="selected"></option>
<option value="1">1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
</select><br/>
</div><br />
<div>
<input id='submit_button'  value="Submit" type="button" />
</div>
</form>`;
                        body.appendChild(canvas);
                        var button;
                        button=document.getElementById('submit_button');
                        button.addEventListener("click", function(){
                            save_submitform(FormDigestValue)
                        } , false)
                        function save_submitform(FormDigestValue){
							var ServRequest1value = document.getElementById("ServiceRequestNumber").value;
                            var CaseOwner1value = document.getElementById("CaseOwnerID").value;
							var nps = document.getElementById("dropdown_nps").value;
							var emotion = document.getElementById("dropdown_sentiment").value;
							var effort = document.getElementById("dropdown_effort").value
							
							console.log('Entered save_submitform() function')
							console.log('Entered: ', FormDigestValue)
							GM.xmlHttpRequest({
								method: "POST",
								url: "http://Sharepoint Site/_api/web/lists/GetByTitle('Predictive-CX')/items",
								data: JSON.stringify({ '__metadata': { 'type': 'SP.Data.PredictiveCXListItem' }, 'Title': ServRequest1value, 'Owner': CaseOwner1value, 'OData__x0051_1': nps, 'OData__x0051_2': emotion, 'OData__x0051_3': effort}),
								headers: {
									"accept": "application/json;odata=verbose",
									"X-RequestDigest": FormDigestValue,
									"content-type": "application/json;odata=verbose",

								},
								onload: function(response) {
									console.log('Response from SHAREPOINT: ', response.response)
								}
							});
						}
                        var trigger = main_frame.createElement("span");
                        trigger.setAttribute("id","triggerMainForm");
                        trigger.setAttribute("style","visibility:visible;float:right;background:#f5fafc;color:#1b527a;border-style:solid;border-width:1px;border-color:#1b527a;z-index:1100;cursor:pointer;");
                        trigger.appendChild(main_frame.createTextNode("PREDICTIVE-CX"));
                        trigger.addEventListener("click",function() {
                            if (main_frame.getElementById("mainForm").style.visibility == "visible") {
                                main_frame.getElementById("mainForm").style.visibility = "hidden";
                            } else {
                                main_frame.getElementById("mainForm").style.visibility = "visible";
                            }
                        });
                        document.getElementById("caseHeader").appendChild(trigger);
                    }
                });
            }
            getFormDigest()

        })();
})();
