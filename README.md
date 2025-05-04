<p align="center">
<img src="https://github.com/aestrus219/sheetflow/assets/122138632/52a01be9-a9a4-43bb-85b0-63cebe401c1d"/>
  <br>
    <br>
<img height="40" src="https://img.shields.io/badge/JavaScript-323330?style=for-the-badge&logo=javascript&logoColor=F7DF1E"/>
<a href="https://www.npmjs.com/package/sheetflow"><img height="40" src="https://img.shields.io/badge/npm-CB3837?style=for-the-badge&logo=npm&logoColor=white"/></a>
<a href="https://github.com/aestrus219/sheetflow/"><img height="40" src="https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white"/></a>
</p>

A lightweight package that allows you to use spreadsheet as a database very easily

# Features
- Easy-to-use methods
- Write, read, update and delete data from spreadsheet
- Search for specific queries in the sheet
- Support for multiple sheets in one doc

# Install
- npm:
  ```
  npm i sheetflow
  ```
- yarn:
  ```
  yarn add sheetflow
  ```

# Basic Usage

- [Example File](https://github.com/vaidik219/sheetflow/blob/main/test/App.test.js)
- [Sample Doc](https://docs.google.com/spreadsheets/d/1SmVEJqgPVDu6_0TyusWpAX7xWV7c6GGrTyv72aY38ac/edit?usp=sharing)

# Guide

## Constructor Options

- ### credentials

  The service account key file.

  - Create a project at [https://console.cloud.google.com](https://console.cloud.google.com)
  - Enable Sheets API. Go to `APIs and services` > `Library`, search and enable `Google Sheets API`
  - Create a service account. Go to `Credentials` > `Create Credentials` > `Service Account`
  - Select the newly created service account, go to `Keys` > `Add new key` > `Create new key` > Select JSON as type and save the JSON file at your project directory

- ### sheet
  The base google sheet.
  - #### id
    The Google doc ID, obtain it from the url :-
    https://docs.google.com/spreadsheets/d/1SmVEJqgPVDu6_0TyusWpAX7xWV7c6GGrTyv72aY38ac/edit?usp=sharing
    <br>
    id for the above sheet: `1SmVEJqgPVDu6_0TyusWpAX7xWV7c6GGrTyv72aY38ac`
  - #### name
    The sheet name from the google doc, you can have multiple class instances for each sheet.
  - #### range
    The range of the stored data, if you don't know what this is just put `A:Z`

## Methods
- ### `db.read()`
  Read all the data from the current sheet - no parameters needed.
- ### `db.write([data])`
  Write to the sheet, param data must be an array of strings.
- ### `db.update(data)`
  Update a cell's value. The param data must be an object, sheetflow by default searches and updates the first entry that has the same value as the first prop of object data.
  ```javascript
  const data = {
    id: "44",
    email: "cool@gmail.com",
  };
  db.update(data); // sets the email of entry by id "44" to cool@gmail.com
  ```
- ### `db.search(query)`
  Search for a specific entry from the database, param query must be a string.
- ### `db.delete(data, all)`
  Deletes the entry by the provided data.
  - #### data
    The param data must be an object, sheetflow by default searches and updates the first entry that has the same value as the first prop of object data.
  - #### all
    Required boolean. Whether to delete all matching entries or not.
    ```javascript
    const data = {
      id: "44",
      email: "cool@gmail.com",
    };
    db.delete(data, true); // deletes all entries with id "44" and email "cool@gmail.com"
    ```
