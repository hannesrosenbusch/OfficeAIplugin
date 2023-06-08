/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { env } from "process";

/* global document, Office, Word */
/* import { openaiApiKey } from './env.js';
 */
async function getResponse(prompt) {
  const { Configuration, OpenAIApi } = require("openai");  
  const apiKey = document.getElementById('userkey').value;
  const configuration = new Configuration({
    apiKey: apiKey,
  });
  const openai = new OpenAIApi(configuration);
  const completion = await openai.createCompletion({
      model: "text-davinci-003",
      prompt: prompt,
      max_tokens: 2048,
    });
  const output = completion.data.choices[0].text.replace(/(\r\n|\n|\r)/gm, ""); //remove linebreaks
  return output.replace(/(\r\n|\n|\r|['"])/gm, ""); //remove quotation marks
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("shorten").onclick = shorten;
    document.getElementById("correct").onclick = correct;
    document.getElementById("paragrSumm").onclick = paragrSummary;
    document.getElementById("custm").onclick = customEdit;
    document.getElementById("custmChat").onclick = chatResponse;

  }
});





export async function shorten() {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    context.load(range, "text");
    await context.sync()
    const selectedText = range.text;
    const newText = await getResponse("Please cut words that are not really necessary from '" + selectedText + "':");
    range.insertText(newText, Word.InsertLocation.replace);
    await context.sync();
  });
}

export async function correct() {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    context.load(range, "text");
    await context.sync()
    const selectedText = range.text;
    const newText = await getResponse("Please correct the English in the following text '" + selectedText + "':");
    range.insertText(newText, Word.InsertLocation.replace);
    await context.sync();
  });
}


export async function paragrSummary() {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    context.load(range, "text");
    await context.sync()
    const selectedText = range.text;
    const ps = range.paragraphs;
    context.load(ps, "items");
    await context.sync();
    let psItems = ps.items;

    for (let i = 0; i < psItems.length; i++) {
      const paragraph = psItems[i];
      context.load(paragraph, "text");
      await context.sync();
      const response = await getResponse("Please write the main message of the following text in one very short sentence '" + paragraph.text + "':");

    paragraph.insertText("\nPARAGRAPH: " + response + "\n", Word.InsertLocation.start);
    await context.sync();
    }
  });
}


export async function customEdit() {
  const textareaValue = document.getElementById('myTextarea').value;
  return Word.run(async (context) => {
    
    const range = context.document.getSelection();
    context.load(range, "text");
    await context.sync()
    const selectedText = range.text;
    const newText = await getResponse("Please apply this edit '" + textareaValue + "' to this text '" + selectedText + "':");
    range.insertText(newText, Word.InsertLocation.replace);
    await context.sync();
  });
}

export async function chatResponse() {
  const { Configuration, OpenAIApi } = require("openai");  
  const apiKey = document.getElementById('userkey').value;
  const configuration = new Configuration({
    apiKey: apiKey,
  });
  const openai = new OpenAIApi(configuration);
  const chatInput = document.getElementById('chatbotInput').value;
    const completion = await openai.createChatCompletion({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: chatInput }],
    });

    const completion_text = completion.data.choices[0].message.content;
    const outputField = document.getElementById('chatbotOutput');
    outputField.value = completion_text;
}








document.addEventListener("DOMContentLoaded", function() {
  var collapsibleBtn = document.querySelector(".collapsible-btn");
  var content = document.querySelector(".content");

  collapsibleBtn.addEventListener("click", function() {
    content.style.display = content.style.display === "none" ? "block" : "none";
  });
});










