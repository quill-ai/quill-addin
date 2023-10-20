/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import axios from "axios";

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});


export async function run() {
  try {
    console.log('trying')
    await Excel.run(async (context) => {
      const descriptionValue = (document.getElementById("descriptionInput")).value;
      console.log(descriptionValue); // This will log the value of the input field

      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      //range.load("address");

      // range.format.fill.color = "red";
      // await context.sync();

      // const options = {
      //   method: 'POST',
      //   headers: {'Content-Type': 'application/json', 'User-Agent': 'insomnia/8.3.0'},
      //   body: '{"name":"test","salary":"123","age":"23"}'
      // };
      
      // fetch('https://dummy.restapiexample.com/api/v1/create', options)
      //   .then(response => response.json())
      //   .then(response => console.log(response))
      //   .catch(err => console.error(err));

      // await fetch("http://localhost:5173/api/formulas", {
      //   // mode: 'cors',
      //   method: "POST",
      //   headers: {
      //     "Content-Type": "application/json"
      //   },
      //   body: JSON.stringify({
      //     "provider": "CIQ",
      //     "prompt": descriptionValue
      //   })
      // })

      // const options = {
      //   method: 'POST',
      //   headers: {'Content-Type': 'application/json'},
      //   body: descriptionValue
      // };

      // const res = await fetch('http://localhost:5173/api/formulas', options)
      //   .then(response => response.json())


      const options = {
        mode: 'cors',
        method: 'POST',
        url: 'http://localhost:5173/api/formulas',
        headers: {'Content-Type': 'application/json'},
        data: {
          provider: 'CIQ',
          prompt: 'Write a formula for TSLAs 2022 calendar gross income'
        }
      };

      const res = await axios.request(options).then(function (response) {
        console.log(response.data);
      })
    

      range.values = [[ 'finished fetch' ]]
      await context.sync();
            
      // const formula = res.formula.toString()

      // range.formulas = [[ formula ]];

      // // Update the fill color
      // range.format.fill.color = "yellow";

      await context.sync();
      // console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.values = [[ error.toString() ]]
      await context.sync();
    });

  }
  // await Excel.run(async (context) => {
  //   const range = context.workbook.getSelectedRange();
  //   range.values = [[ 'exited' ]]
  //   await context.sync();
  // });
}
