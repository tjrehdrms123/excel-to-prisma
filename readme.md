# excel-to-prisma

`excel-to-prisma` is a package that parses Excel data and converts it into a data object that can be used in Prisma ORM.

## Installation

```bash
npm i excel-to-prisma
```

## Usage

excel-to-prisma works as if it reads the parent table and concatenates the child tables.

### Initializing excel-to-prisma

First, import and create an instance of excel-to-prisma with the path to your Excel file and the worksheet name you intend to work with:

```js
import { ExcelToPrisma } from "excel-to-prisma";

const excelToPrisma = new ExcelToPrisma({
  filePath: "./data.xlsx", // your xlsx file
  oneToOneOrManyConnectOptions: {
    keyword: "info", // Separator string for one-to-many relationship
    split: "|", // String to separate multiple values
  },
});
await excelToPrisma.initialize();
```

### Read parent table

When parsing the parent table, write it as follows:

```js
await excelToPrisma.readSheet({
  name: "user",
  rowNameIndex: 2,
  startRowIndex: 3,
});
```

### One to Many Create

When linking to a parent table in a one to many relationship, you would write:

The oneToManyCreate method is executing the oneToOneOrManyConnect method

```js
await excelToPrisma.oneToManyCreate({
  name: "post",
  fk: "userId",
  rowNameIndex: 2,
  startRowIndex: 3,
});
```

### Subtables in relationship

When connecting child tables in relationship, write as follows:

```js
await excelToPrisma
  .readSheet({ name: "user", rowNameIndex: 2, startRowIndex: 3 })
  .then(async (sheetOption) => {
    await excelToPrisma
      .oneToManyCreate({
        name: "product",
        fk: "userId",
        many: sheetOption.name,
        rowNameIndex: 2,
        startRowIndex: 3,
      })
      .then(async (sheetOption) => {
        await excelToPrisma.oneToManyCreate({
          name: "productComment",
          fk: "productId",
          many: sheetOption.name,
          rowNameIndex: 2,
          startRowIndex: 3,
        });
      });
  });
```

### Axios usage and example code

```js
import { ExcelToPrisma } from "excel-to-prisma";
import axios from "axios";

async function main() {
  // parse excel to prisma
  const excelToPrisma = new ExcelToPrisma({
    filePath: "./data.xlsx",
    oneToOneOrManyOptions: {
      keyword: "info",
      split: "|",
    },
  });
  await excelToPrisma.initialize();
  await excelToPrisma
    .readSheet({ name: "user", rowNameIndex: 2, startRowIndex: 3 })
    .then(async (sheetOption) => {
      await excelToPrisma.oneToManyCreate({
        name: "post",
        fk: "userId",
        many: sheetOption.name,
        rowNameIndex: 2,
        startRowIndex: 3,
      });
      await excelToPrisma
        .oneToManyCreate({
          name: "product",
          fk: "userId",
          many: sheetOption.name,
          rowNameIndex: 2,
          startRowIndex: 3,
        })
        .then(async (sheetOption) => {
          await excelToPrisma
            .oneToManyCreate({
              name: "productComment",
              fk: "productId",
              many: sheetOption.name,
              rowNameIndex: 2,
              startRowIndex: 3,
            })
            .then(async (sheetOption) => {
              await excelToPrisma.oneToManyCreate({
                name: "productCommentHistory",
                fk: "productCommentId",
                many: sheetOption.name,
                rowNameIndex: 2,
                startRowIndex: 3,
              });
            });
        });
    });
  const data = JSON.stringify(excelToPrisma.getData()); // stringified data

  // axios post to prisma
  for (let i = 0; i < JSON.parse(data).length; i++) {
    await axios
      .post("homepage URL", JSON.parse(data)[i], {
        headers: {
          "Content-Type": "application/json",
        },
      })
      .then((res) => {
        console.log(res.data);
      })
      .catch((err) => {
        console.error(err);
      });
  }
}

main().catch((err) => console.error(err));
```

## API Reference

Refer to the code comments for detailed API usage and method descriptions.

## Example

- [Excel](https://github.com/tjrehdrms123/excel-to-prisma/tree/main/src/assets/data.xlsx)

- [Test Code](https://github.com/tjrehdrms123/excel-to-prisma/tree/main/src/tests/base.spec.ts)

- [Output JSON](https://github.com/tjrehdrms123/excel-to-prisma/tree/main/src/assets/output.json)

## Contact

Seog Donggeun - seogdonggeun@gmail.com

Project Link: [excel-to-prisma](https://github.com/tjrehdrms123/excel-to-prisma)
