//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\///\\//\\//\\//\\//\\//
//   Hard coded Leaese Stockboard                                \\
//   Obrien 2026-PrintCo_8x3 Stockboard_P                        \\
// Form: Obrien 2026 - PrintCo Stockboard (8x3) No SaleM         \\
//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\///\\//\\//\\//\\//\\//


var agent1 = "";
var agent2 = "";
var salemethod = "For Sale";
var name1Array = Artwork.pv("AC_AgentName1").split(" ");
var a1FirstName = name1Array.slice(0, -1).join(" ");
var a1LastName = name1Array[name1Array.length - 1];

var name2Array = Artwork.pv("AC_AgentName2").split(" ");
var a2FirstName = name2Array.slice(0, -1).join(" ");
var a2LastName = name2Array[name1Array.length - 1];

Artwork.SetImagePOSourceImage("AHS1", -2147483648);
Artwork.SetImagePOSourceImage("AHS2", -2147483648);
Artwork.SetTextBoxContents("AgentInfoNHS", "");
Artwork.SetTextBoxContents("Agent1Info", "");
Artwork.SetTextBoxContents("Agent2Info", "");
Artwork.SetTextBoxContents("Address", "");
Artwork.SetTextBoxContents("URL", "");

// Sale Method (hard coded to For Sale For now)
// var saleType = Artwork.pv("AC_MultiuseDropdown2");
// var saleTypeLabels = {
//   "For Sale": "For<SCR>Sale",
//   "For Lease": "For<SCR>Lease",
// };


//\\//\\//\\//\\//\\//\\//\\//
//     SIGNBOARD  STYLE     \\
//\\//\\//\\//\\//\\//\\//\\//

// FOREST
if (Artwork.pv("AC_MultiuseDropdown") === "Forest") {
    Artwork.SetPOPos("bgForest", "0", "0"); // Move Forest Background onto page
    Artwork.SetPOPos("bgLinen", "800", "0"); // Move Linen Background off page
    Artwork.SetImagePOSourceImage("Logo", -2140266962);
    textcol = "Linen"; // Set text colour
    qrcol = "l";

// LINEN
} else if (Artwork.pv("AC_MultiuseDropdown") === "Linen") {
    Artwork.SetPOPos("bgLinen", "0", "0"); // Move Linen Background onto page
    Artwork.SetPOPos("bgForest", "800", "0"); // Move Forest Background off page
    Artwork.SetImagePOSourceImage("Logo", -2140266964);
    textcol = "Forest"; // Set text colour
    qrcol = "f";
}

// QR Code
if (Artwork.pv("AC_WebURL") !== "") {
    Artwork.SetImagePOSourceImageByQueueAndName("QR1", -2146851632, Artwork.pv("AC_ArtworkID") + qrcol); //White QR Code for Charcoal Brochure
    Artwork.SetTextBoxContents("URL", "[Website" + textcol + "]obre.com.au");
} else {
    Artwork.SetImagePOSourceImage("QR1", -2147483648);
    Artwork.SetTextBoxContents("URL", "");
}


//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//
//       Sale Method                        \\ 
// Left optional commented in in case       \\
//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//
// if (saleTypeLabels[saleType] !== "undefined") {
//     salemethod += "[SaleMethodMoss]" + saleTypeLabels[saleType];
//     Artwork.SetTextBoxContents("SaleMethod", salemethod);
// }
 
   Artwork.SetTextBoxContents("SaleMethod", "[SaleMethodMoss]For<SCR>Sale");

//\\//\\//\\//\\//\\//\\//\\//
//      Agent Details       \\
//\\//\\//\\//\\//\\//\\//\\//
if (Artwork.pv("AC_AgentName1") !== "") {
    if (Artwork.pv("AC_AgentName1").length <= 20 || Artwork.pv("AC_CopySix") === "" || salemethod !== "For Sale") {
        agent1 += "[TextReg" + textcol + "]" + Artwork.pv("AC_AgentName1") + "<SCR>";
    } else {
        agent1 += "[TextReg" + textcol + "]" + a1FirstName  + "<SCR>";
        agent1 += "[TextReg" + textcol + "]" + a1LastName  + "<SCR>";
    }
    
    agent1 += "[TextReg" + textcol + "]" + Artwork.pv("AC_AgentPhone1");
    
    if (Artwork.pv("AC_AgentName2") !== "") {
        if (Artwork.pv("AC_AgentName2").length <= 20 || Artwork.pv("AC_CopySix") === "" || salemethod !== "For Sale") {
            agent2 += "[TextReg" + textcol + "]" + Artwork.pv("AC_AgentName2") + "<SCR>";
        } else {
            agent2 += "[TextReg" + textcol + "]" + a2FirstName  + "<SCR>";
            agent2 += "[TextReg" + textcol + "]" + a2LastName + "<SCR>";
        }
        
        agent2 += "[TextReg" + textcol + "]" + Artwork.pv("AC_AgentPhone2");
        
    }
}


if (salemethod === "For Sale" && Artwork.pv("AC_CopySix") !== "") {
    Artwork.SetImagePOSourceImageByQueueAndName("AHS1", -2146768183, Artwork.pvf("AC_AgentName1", "cleanText") + "_Round.png");
    Artwork.SetImagePOSourceImageByQueueAndName("AHS2", -2146768183, Artwork.pvf("AC_AgentName2", "cleanText") + "_Round.png");
    Artwork.SetTextBoxContents("Agent1Info", agent1);
    Artwork.SetTextBoxContents("Agent2Info", agent2);
} else {
    agent1 = agent1 + "[TextReg" + textcol + "] <CR>" + agent2;
    Artwork.SetTextBoxContents("AgentInfoNHS", agent1);
}


//<Cleanup>
Artwork.PlaceImages();
Artwork.UpdatePOTextBoxOverSetInformation();
//</Cleanup>
