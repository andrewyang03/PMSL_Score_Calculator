        let xlsx=require("xlsx"); 
        let workBook=xlsx.readFile("Divisional.xlsx"); 

        let dataTally= {}; 
        for (const sheetName of workBook.SheetNames){
            dataTally[sheetName]=xlsx.utils.sheet_to_json(workBook.Sheets[sheetName]); 
        }
        //let xlsx=require("xlsx"); 
        //const jsontoxml=require("jsontoxml"); 


        
        //console.log("json: \n", JSON.stringify(dataTally), "\n \n"); 

        let cheverly_Total=0; 
        let seals_Total=0; 
        let sbpDolphins_Total=0; 
        let mvpDolphins_Total=0; 
        let bsrPM_Total=0; 
        let westArundel_Total=0; 
        //console.log(JSON.stringify(dataTally["#LN00294"].filter(k=>k.Event===43))); 
        for (let i=3; i<47; i++){
            //let eventRanking=dataTally["#LN00294"].filter(e=> e.Event===i); 
            let eventRanking=sortTimes(dataTally["#LN00294"].filter(e=> e.Event===i)); 
            if (eventRanking[0].Team==="Cheverly"){
                cheverly_Total+=7; 
            }
            if (eventRanking[0].Team==="Seals"){
                seals_Total+=7; 
            }
            if (eventRanking[0].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=7; 
            }
            if (eventRanking[0].Team==="MVP_B"){
                mvpDolphins_Total+=7; 
            }
            if (eventRanking[0].Team==="BSR-PM"){
                bsrPM_Total+=7; 
            }
            if (eventRanking[0].Team==="WA"){
                westArundel_Total+=7; 
            }
            if (eventRanking[1].Team==="Cheverly"){
                cheverly_Total+=5; 
            }
            if (eventRanking[1].Team==="Seals"){
                seals_Total+=5; 
            }
            if (eventRanking[1].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=5; 
            }
            if (eventRanking[1].Team==="MVP_B"){
                mvpDolphins_Total+=5; 
            }
            if (eventRanking[1].Team==="BSR-PM"){
                bsrPM_Total+=5; 
            }
            if (eventRanking[1].Team==="WA"){
                westArundel_Total+=5; 
            }
            if (eventRanking[2].Team==="Cheverly"){
                cheverly_Total+=4; 
            }
            if (eventRanking[2].Team==="Seals"){
                seals_Total+=4; 
            }
            if (eventRanking[2].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=4; 
            }
            if (eventRanking[2].Team==="MVP_B"){
                mvpDolphins_Total+=4; 
            }
            if (eventRanking[2].Team==="BSR-PM"){
                bsrPM_Total+=4; 
            }
            if (eventRanking[2].Team==="WA"){
                westArundel_Total+=4; 
            }
            if (eventRanking[3].Team==="Cheverly"){
                cheverly_Total+=3; 
            }
            if (eventRanking[3].Team==="Seals"){
                seals_Total+=3; 
            }
            if (eventRanking[3].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=3; 
            }
            if (eventRanking[3].Team==="MVP_B"){
                mvpDolphins_Total+=3; 
            }
            if (eventRanking[3].Team==="BSR-PM"){
                bsrPM_Total+=3; 
            }
            if (eventRanking[3].Team==="WA"){
                westArundel_Total+=3; 
            }
            if (eventRanking[4].Team==="Cheverly"){
                cheverly_Total+=2; 
            }
            if (eventRanking[4].Team==="Seals"){
                seals_Total+=2; 
            }
            if (eventRanking[4].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=2; 
            }
            if (eventRanking[4].Team==="MVP_B"){
                mvpDolphins_Total+=2; 
            }
            if (eventRanking[4].Team==="BSR-PM"){
                bsrPM_Total+=2; 
            }
            if (eventRanking[4].Team==="WA"){
                westArundel_Total+=2; 
            }
            if (eventRanking[5].Team==="Cheverly"){
                cheverly_Total+=1; 
            }
            if (eventRanking[5].Team==="Seals"){
                seals_Total+=1; 
            }
            if (eventRanking[5].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=1; 
            }
            if (eventRanking[5].Team==="MVP_B"){
                mvpDolphins_Total+=1; 
            }
            if (eventRanking[5].Team==="BSR-PM"){
                bsrPM_Total+=1; 
            }
            if (eventRanking[5].Team==="WA"){
                westArundel_Total+=1; 
            }

        }
        for (let i=1; i<=2; i++){
            let relayRanking_Medley=sortTimes(dataTally["#LN00294"].filter(e=> e.Event===i)); 
            if (relayRanking_Medley[0].Team==="Cheverly"){
                cheverly_Total+=14; 
            }
            if (relayRanking_Medley[0].Team==="Seals"){
                seals_Total+=14; 
            }
            if (relayRanking_Medley[0].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=14; 
            }
            if (relayRanking_Medley[0].Team==="MVP_B"){
                mvpDolphins_Total+=14; 
            }
            if (relayRanking_Medley[0].Team==="BSR-PM"){
                bsrPM_Total+=14; 
            }
            if (relayRanking_Medley[0].Team==="WA"){
                westArundel_Total+=14; 
            }
            if (relayRanking_Medley[1].Team==="Cheverly"){
                cheverly_Total+=10; 
            }
            if (relayRanking_Medley[1].Team==="Seals"){
                seals_Total+=10; 
            }
            if (relayRanking_Medley[1].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=10; 
            }
            if (relayRanking_Medley[1].Team==="MVP_B"){
                mvpDolphins_Total+=10; 
            }
            if (relayRanking_Medley[1].Team==="BSR-PM"){
                bsrPM_Total+=10; 
            }
            if (relayRanking_Medley[1].Team==="WA"){
                westArundel_Total+=10; 
            }
            if (relayRanking_Medley[2].Team==="Cheverly"){
                cheverly_Total+=8; 
            }
            if (relayRanking_Medley[2].Team==="Seals"){
                seals_Total+=8; 
            }
            if (relayRanking_Medley[2].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=8; 
            }
            if (relayRanking_Medley[2].Team==="MVP_B"){
                mvpDolphins_Total+=8; 
            }
            if (relayRanking_Medley[2].Team==="BSR-PM"){
                bsrPM_Total+=8; 
            }
            if (relayRanking_Medley[2].Team==="WA"){
                westArundel_Total+=8; 
            }
            if (relayRanking_Medley[3].Team==="Cheverly"){
                cheverly_Total+=6; 
            }
            if (relayRanking_Medley[3].Team==="Seals"){
                seals_Total+=6; 
            }
            if (relayRanking_Medley[3].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=6; 
            }
            if (relayRanking_Medley[3].Team==="MVP_B"){
                mvpDolphins_Total+=6; 
            }
            if (relayRanking_Medley[3].Team==="BSR-PM"){
                bsrPM_Total+=6; 
            }
            if (relayRanking_Medley[3].Team==="WA"){
                westArundel_Total+=6; 
            }
            if (relayRanking_Medley[4].Team==="Cheverly"){
                cheverly_Total+=4; 
            }
            if (relayRanking_Medley[4].Team==="Seals"){
                seals_Total+=4; 
            }
            if (relayRanking_Medley[4].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=4; 
            }
            if (relayRanking_Medley[4].Team==="MVP_B"){
                mvpDolphins_Total+=4; 
            }
            if (relayRanking_Medley[4].Team==="BSR-PM"){
                bsrPM_Total+=4; 
            }
            if (relayRanking_Medley[4].Team==="WA"){
                westArundel_Total+=4; 
            }
            if (relayRanking_Medley[5].Team==="Cheverly"){
                cheverly_Total+=2; 
            }
            if (relayRanking_Medley[5].Team==="Seals"){
                seals_Total+=2; 
            }
            if (relayRanking_Medley[5].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=2; 
            }
            if (relayRanking_Medley[5].Team==="MVP_B"){
                mvpDolphins_Total+=2; 
            }
            if (relayRanking_Medley[5].Team==="BSR-PM"){
                bsrPM_Total+=2; 
            }
            if (relayRanking_Medley[5].Team==="WA"){
                westArundel_Total+=2; 
            }
        }
        for (let i=47; i<=49; i++){
            let relayRanking_Free=sortTimes(dataTally["#LN00294"].filter(e=> e.Event===i)); 
            if (relayRanking_Free[0].Team==="Cheverly"){
                cheverly_Total+=14; 
            }
            if (relayRanking_Free[0].Team==="Seals"){
                seals_Total+=14; 
            }
            if (relayRanking_Free[0].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=14; 
            }
            if (relayRanking_Free[0].Team==="MVP_B"){
                mvpDolphins_Total+=14; 
            }
            if (relayRanking_Free[0].Team==="BSR-PM"){
                bsrPM_Total+=14; 
            }
            if (relayRanking_Free[0].Team==="WA"){
                westArundel_Total+=14; 
            }
            if (relayRanking_Free[1].Team==="Cheverly"){
                cheverly_Total+=10; 
            }
            if (relayRanking_Free[1].Team==="Seals"){
                seals_Total+=10; 
            }
            if (relayRanking_Free[1].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=10; 
            }
            if (relayRanking_Free[1].Team==="MVP_B"){
                mvpDolphins_Total+=10; 
            }
            if (relayRanking_Free[1].Team==="BSR-PM"){
                bsrPM_Total+=10; 
            }
            if (relayRanking_Free[1].Team==="WA"){
                westArundel_Total+=10; 
            }
            if (relayRanking_Free[2].Team==="Cheverly"){
                cheverly_Total+=8; 
            }
            if (relayRanking_Free[2].Team==="Seals"){
                seals_Total+=8; 
            }
            if (relayRanking_Free[2].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=8; 
            }
            if (relayRanking_Free[2].Team==="MVP_B"){
                mvpDolphins_Total+=8; 
            }
            if (relayRanking_Free[2].Team==="BSR-PM"){
                bsrPM_Total+=8; 
            }
            if (relayRanking_Free[2].Team==="WA"){
                westArundel_Total+=8; 
            }
            if (relayRanking_Free[3].Team==="Cheverly"){
                cheverly_Total+=6; 
            }
            if (relayRanking_Free[3].Team==="Seals"){
                seals_Total+=6; 
            }
            if (relayRanking_Free[3].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=6; 
            }
            if (relayRanking_Free[3].Team==="MVP_B"){
                mvpDolphins_Total+=6; 
            }
            if (relayRanking_Free[3].Team==="BSR-PM"){
                bsrPM_Total+=6; 
            }
            if (relayRanking_Free[3].Team==="WA"){
                westArundel_Total+=6; 
            }
            if (relayRanking_Free[4].Team==="Cheverly"){
                cheverly_Total+=4; 
            }
            if (relayRanking_Free[4].Team==="Seals"){
                seals_Total+=4; 
            }
            if (relayRanking_Free[4].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=4; 
            }
            if (relayRanking_Free[4].Team==="MVP_B"){
                mvpDolphins_Total+=4; 
            }
            if (relayRanking_Free[4].Team==="BSR-PM"){
                bsrPM_Total+=4; 
            }
            if (relayRanking_Free[4].Team==="WA"){
                westArundel_Total+=4; 
            }
            if (relayRanking_Free[5].Team==="Cheverly"){
                cheverly_Total+=2; 
            }
            if (relayRanking_Free[5].Team==="Seals"){
                seals_Total+=2; 
            }
            if (relayRanking_Free[5].Team==="Sbp Dolphins"){
                sbpDolphins_Total+=2; 
            }
            if (relayRanking_Free[5].Team==="MVP_B"){
                mvpDolphins_Total+=2; 
            }
            if (relayRanking_Free[5].Team==="BSR-PM"){
                bsrPM_Total+=2; 
            }
            if (relayRanking_Free[5].Team==="WA"){
                westArundel_Total+=2; 
            }
        }


        console.log("Roger Carter Seals: ", seals_Total); 
        console.log("Cheverly Dolphins: ", cheverly_Total); 
        console.log("MVP Dolphins: ", mvpDolphins_Total); 
        console.log("Strathmore Bel Pre Dolphins: ", sbpDolphins_Total); 
        console.log("Belair Swim & Raquet Orange Barracudas: ", bsrPM_Total); 
        console.log("West Arundel Aqua Ducks: ", westArundel_Total); 
        function sortTimes(data){
            return data.sort((a, b)=> parseFloat(a["Best Time"])-parseFloat(b["Best Time"])); 
        }