const { Configuration, OpenAIApi } = require("openai");
const xlsx = require('xlsx');
const axios = require('axios');
const cheerio = require('cheerio');
const puppeteer = require('puppeteer');
const fs = require('fs');

const http = require('http');
const https = require('https');

const configuration = new Configuration({
  apiKey: "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",
});




// set the URL you want to retrieve

// (async()=>{
//     const url = 'https://www.aliexpress.com/item/1005004737532420.html?srcSns=sns_Copy&spreadType=socialShare&bizType=ProductDetail&social_params=20644166989&aff_fcid=4582164264bd4faca519d617b17805b2-1681130059921-06246-_mMv2ITS&tt=MG&aff_fsk=_mMv2ITS&aff_platform=default&sk=_mMv2ITS&aff_trace_key=4582164264bd4faca519d617b17805b2-1681130059921-06246-_mMv2ITS&shareId=20644166989&businessType=ProductDetail&platform=AE&terminal_id=5351ab3bb8af4cd387f2f898d95a394f&afSmartRedirect=y';
//     const html = await fetchHtml(url);
//     fs.writeFileSync('example.html', html);
//     console.log(html);
// })()
// async function fetchHtml(url) {
//   const browser = await puppeteer.launch();
//   const page = await browser.newPage();
//   await page.goto(url, { waitUntil: 'networkidle2' });
//   const html = await page.content();
//   await browser.close();
//   return html;
// }






let datatopush=[];
async function getAiResponse(topic,url) {
    console.log('topic>>>',topic)

    if(topic=='404'){
        console.log("404 error")

 // Load existing workbook
 datatopush.push({Notes:url, Description:"not found",Title: "not found",Summeries:"not found"  })

    }
    else{

  const openai = new OpenAIApi(configuration);
  const description = await openai.createCompletion({
    model: "text-davinci-003",
    prompt: topic +"\n"+ "rewrite the above description to 500-1800 characters. The description must include selling keywords used in google trends for the set product. The description is not allowed to exceed the number of characters. The description should have bold sections for the Item description, technical specification, and included in package. The description should use bullet points where it is suitable. Do this with HTML.",
    max_tokens: 1024,
    n: 1,
    stop: null,
    temperature: 0.7
  });

  const title = await openai.createCompletion({
    model: "text-davinci-003",
    prompt: topic +"\n"+"from the above description generate a selling title which is suitable for the product The title is not allowed to exceed 44 characters.",
    max_tokens: 1024,
    n: 1,
    stop: null,
    temperature: 0.7
  });

  const sellingSummaries = await openai.createCompletion({
    model: "text-davinci-003",
    prompt: topic +"\n"+"from the above description generate five unique selling summaries with keywords, The summaries must be unique. The summaries are not allowed to exceed 44 characters. Do this with HTML.",
    max_tokens: 1024,
    n: 1,
    stop: null,
    temperature: 0.7
  });

  let descriptiongen=description.data.choices[0].text;
  let titlegen=title.data.choices[0].text;
  let summeriesgen=sellingSummaries.data.choices[0].text;
  console.log("chatgptreturned description",description.data.choices[0].text);
  console.log("chatgptreturned title",title.data.choices[0].text);
  console.log("chatgptreturned sellingSummaries",sellingSummaries.data.choices[0].text);
  datatopush.push({Notes:url, Description:descriptiongen,Title: titlegen,Summeries:summeriesgen  })
}



// Define data to write to file


}

const pushtocsv=()=>{
    const workbook1= xlsx.utils.book_new();
    const sheet = xlsx.utils.json_to_sheet(datatopush);
    xlsx.utils.book_append_sheet(workbook1, sheet, 'Sheet1');
    xlsx.writeFile(workbook1, 'example.xlsx');
}


//getAiResponse("Hello, can you help me write a poem?");

// open excel sheet and scrap first column

const workbook = xlsx.readFile('export_2023-04-08T01 08 22.717Z.xlsx');
const sheet_name = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheet_name];

const column_name = 'A';
const range = worksheet['!ref'];

const rows = range.split(':');
const start_row = parseInt(rows[0].match(/\d+/));
const end_row = parseInt(rows[1].match(/\d+/));
console.log("endrow>>>",end_row)

async function processUrls() {

for (let i = start_row+1; i <= 2; i++) {   //end_row
  const cell_address = column_name + i.toString();
  const cell = worksheet[cell_address];
  if (cell) {
    console.log("url>>>>>>>>>>>>>>>>>>>>>>>>>.","'"+cell.v+"'");
    
//enter link to be scrapped

    //const url = cell.v;
    const url ="https://www.aliexpress.com/item/1005004737532420.html?srcSns=sns_Copy&spreadType=socialShare&bizType=ProductDetail&social_params=20644166989&aff_fcid=4582164264bd4faca519d617b17805b2-1681130059921-06246-_mMv2ITS&tt=MG&aff_fsk=_mMv2ITS&aff_platform=default&sk=_mMv2ITS&aff_trace_key=4582164264bd4faca519d617b17805b2-1681130059921-06246-_mMv2ITS&shareId=20644166989&businessType=ProductDetail&platform=AE&terminal_id=5351ab3bb8af4cd387f2f898d95a394f&afSmartRedirect=y"

async function getTitle(url) {
    var productDescription;
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    
    try {
      
      await page.goto(url,{ waitUntil: 'load', timeout: 0 });
      
      // Wait for the element to appear on the page
      await page.waitForSelector('.product-description');
      console.log("it waited for the id")
      // Extracting text from the div with id `product-description`
      productDescription = await page.$eval('.product-description', element => element.textContent.trim());
      console.log(productDescription);
      await getAiResponse(productDescription,url);
    } catch (error) {
      console.error(`Error loading page: ${error}`);
      productDescription="404"
     await getAiResponse(productDescription,url);
      // Handle the error here, for example:
      // throw new Error('Page not found');
    } finally {
      await browser.close();
    }
  }
  
  
  await getTitle(url)
    .then(title => console.log(title))
    .catch(error => console.error(error));

  }

 
}
pushtocsv()
}
processUrls();