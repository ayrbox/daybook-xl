#!/usr/bin/env node

const arg = require("arg");
const inquirer = require("inquirer");

const createDaybook = require("../src/index");

function parseArgsIntoOptions(rawArgs) {
  const args = arg(
    {
      "--file-name": String,
      "-f": "--file-name"
    },
    {
      argv: rawArgs.slice(2)
    }
  );

  return {
    skipPrompts: !!args["--file-name"] || false,
    filename: args["--file-name"] || ""
  };
}

async function promptForMissingOptions(options) {
  const defaultFilename = "daybook.xlsx";
  if (options.skipPrompts) {
    return {
      ...options,
      filename: options.filename || defaultFilename
    };
  }

  const questions = [];
  if (!options.filename) {
    questions.push({
      type: "input",
      name: "filename",
      message: "Please enter daybook filename to create"
    });
  }

  const answers = await inquirer.prompt(questions);
  return {
    ...options,
    filename: options.filename || answers.filename || defaultFilename
  };
}

const args = parseArgsIntoOptions(process.argv);

promptForMissingOptions(args)
  .then(options => {
    console.log(options);
    createDaybook(options);
  })
  .then(() => console.log("Daybook created."))
  .catch(err => console.log("Unable to create daybook."));
