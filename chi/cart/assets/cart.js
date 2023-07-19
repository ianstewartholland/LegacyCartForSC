
// State lists for canada and usa
var states = new Array();

states['Canada'] = new Array('Please select...','Alberta - AB','British Columbia - BC','Manitoba - MB','New Brunswick - NB','Newfoundland - NL','Northwest Territories - NT','Nova Scotia - NS','Nunavut - NU','Ontario - ON','Prince Edward Island - PE','Quebec - QC','Saskatchewan - SK','Yukon Territory - YT');
states['USA'] = new Array('Please select...','Alabama - AL','Alaska - AK','Arizona - AZ','Arkansas - AR','California - CA','Colorado - CO','Connecticut - CT','Delaware - DE','District of Columbia - DC','Florida - FL','Georgia - GA','Hawaii - HI','Idaho - ID','Illinois - IL','Indiana - IN','Iowa - IA','Kansas - KS','Kentucky - KY','Louisiana - LA','Maine - ME','Maryland - MD','Massachusetts - MA','Michigan - MI','Minnesota - MN','Mississippi - MS','Missouri - MO','Montana - MT','Nebraska - NE','Nevada - NV','New Hampshire - NH','New Jersey - NJ','New Mexico - NM','New York - NY','North Carolina - NC','North Dakota - ND','Ohio - OH','Oklahoma - OK','Oregon - OR','Pennsylvania - PA','Rhode Island - RI','South Carolina - SC','South Dakota - SD','Tennessee - TN','Texas - TX','Utah - UT','Vermont - VT','Virginia - VA','Washington - WA','West Virginia - WV','Wisconsin - WI','Wyoming - WY','American Samoa - AS','Guam - GU','Federated States of Micronesia - FM','Marshall Islands - MH','Northern Mariana Islands - MP','Palau - PW','Puerto Rico - PR','Virgin Islands - WY');

function setStates() {
  var cntrySel = document.getElementById('Country');
  var stateList;
  if(cntrySel.selectedIndex == 1)
    stateList = states ['USA'];
  else if (cntrySel.selectedIndex == 2)
    stateList = states['Canada'];
  
  changeSelect('Province', stateList, stateList);
}

function changeSelect(fieldID, newOptions, newValues)
{
  selectField = document.getElementById(fieldID);
  var parentObj;
  
  if(document.getElementById("Country").selectedIndex > 2)
  {
    parentObj = selectField.parentNode;
    parentObj.removeChild(selectField);
    var inputEl = document.createElement("INPUT");
    inputEl.setAttribute("id", "Province");
    inputEl.setAttribute("type", "text");
    inputEl.setAttribute("name", "Province");
    inputEl.setAttribute("size", 20);
    parentObj.appendChild(inputEl) ;
  }
  else
  {
    if (selectField.type == 'text' )
    {
        parentObj = selectField.parentNode;
        parentObj.removeChild(selectField);
        var inputSel = document.createElement("SELECT");
        inputSel.setAttribute("name","Province");
        inputSel.setAttribute("id","Province");
        parentObj.appendChild(inputSel) ;
        selectField = document.getElementById('Province');
    }
   
    selectField.options.length = 0;
    if(newOptions == null)
    {
      selectField.options[0] = new Option("Please select country first");
    }
    for (i=0; i < newOptions.length; i++)
    {
        selectField.options[selectField.length] = new Option(newOptions[i], newValues[i]);
    }  
  }
}

function pageLoad(){
    var selectField = document.getElementById("Province");
    var parentObj;
    if(document.getElementById('Country').selectedIndex == 0){
      parentObj = selectField.parentNode;
      parentObj.removeChild(selectField);
      var inputSel = document.createElement("SELECT");
      inputSel.setAttribute("name","Province");
      inputSel.setAttribute("id","Province");
      parentObj.appendChild(inputSel) ;      
      selectField = document.getElementById('Province');
      selectField.options[0] = new Option("Please select country first");      
    }

    var selectCountry = document.getElementById('Country');
    if(selectCountry.addEventListener )
      selectCountry.addEventListener("change", setStates, false);
    if(selectCountry.attachEvent)
      selectCountry.attachEvent("change", setStates);
}

function validate_select(field,alerttxt) {
    with (field) {
        if (field.selectedIndex == 0 && field.length > 1) {
            alert(alerttxt);
            return false;
        }
        else {
            return true;
        }
    }
}

function validate_required(field,alerttxt) {
    with (field) {
        if (value==null||value=="") {
            alert(alerttxt);
            return false;
        }
        else {
            return true;
        }
    }
}

function validate_form(thisform) {
    with (thisform) {

        var sel = document.getElementById("Country");
        if (sel != null) {
            if (sel.options[sel.selectedIndex].value == "Nigeria" || sel.options[sel.selectedIndex].value == "Cameroon" ||
                sel.options[sel.selectedIndex].value == "Uganda" || sel.options[sel.selectedIndex].value == "Nepal" ||
                sel.options[sel.selectedIndex].value == "Ghana" || sel.options[sel.selectedIndex].value == "Cote d'Ivoire" ||
                sel.options[sel.selectedIndex].value == "Tanzania") {
                alert("If you would like to continue to place an order, please contact CHI.");
                return false;
            }
        }

        if (validate_required(FirstName,"Please enter a first name.")==false) {
            FirstName.focus();
            return false;
            }
            
        else if (validate_required(LastName,"Please enter a last name.")==false) {
            LastName.focus();
            return false;
            }
        else if (validate_required(Company,"Please enter your company name, school or self employed. ")==false) {
            Company.focus();
            return false;
            }
            
        else if (validate_required(StreetAddress,"Please enter your address.")==false) {
            StreetAddress.focus();
            return false;
            }
            
        else if (validate_required(City,"Please enter your city.")==false) {
            City.focus();
            return false;
            }

        else if (validate_select(Country,"Please select a country.")==false) {
            Country.focus();
            return false;
            }
        else if (validate_select(Province,"Please select a province/state.")==false) {
            Province.focus();
            return false;
            }
//        else if (validate_required(Province,"Please enter your province/state/region/district.")==false) {
//            Province.focus();
//            return false;
//            }

        else if (validate_required(PostalCode,"Please enter your postal or zip code.")==false) {
            PostalCode.focus();
            return false;
            }      
        
        else if (validate_required(Email,"A valid email address is required.")==false) {
            Email.focus();
            return false;
            }
            
        else if (validate_required(Telephone,"A telephone number is required so that we may contact with any possible issues regarding your order.")==false) {
            Telephone.focus();
            return false;
        }

        else if (validate_select(WExperience1, "Please enter your experience on hydrology/hydraulic modeling software.") == false) {
            WExperience1.focus();
            return false;
        }
        else if (validate_select(WExperience2, "Please enter your experience on EPA SWMM5 or PCSWMM.") == false) {
            WExperience2.focus();
            return false;
        }
        else if (validate_select(WInterest1, "Please enter what workshop topics are of interest to you.") == false) {
            WInterest1.focus();
            return false;
        }



    }
}

function loadButton() {
  var buttonnode= document.createElement('input');
  buttonnode.setAttribute('type','button');
  buttonnode.setAttribute('name','btnPrint');
  buttonnode.setAttribute('value','Print this page');
  buttonnode.onclick = function ClickToPrint() {
     var WinPrint = window.open('', '', 'left=0,top=0,width=820,height=900');
     var PrintDiv = document.getElementById('printContent').innerHTML;
     WinPrint.document.write('<html><body style ="font-family:Arial, Helvetica, sans-serif; font-size:11px; width:620px; padding-top: 30px; padding-left:30px; padding-right:30px;">');
     WinPrint.document.write(PrintDiv);
     WinPrint.document.write('</body></html>');
     WinPrint.document.close();
     WinPrint.focus();
     WinPrint.print();
     WinPrint.close();
}	    

  var foo = document.getElementById("fooBar");
  foo.appendChild(buttonnode);
}
