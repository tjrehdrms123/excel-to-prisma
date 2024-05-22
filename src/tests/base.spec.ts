import { ExcelToPrisma } from "../excel-to-prisma";

describe('ExcelToPrisma tests', () => {
  let excelToPrisma: ExcelToPrisma;
  beforeEach(async() => {
    excelToPrisma = new ExcelToPrisma({
      filePath: 'src/assets/data.xlsx',
      pkDelimiterString: "Id",
      oneToOneOrManyOptions : {
        split: '|'
      }
    });
    await excelToPrisma.initialize();
  });

  it('should read single table', async () => {
    await excelToPrisma.readSheet({ name: "banner", rowNameIndex: 2, startRowIndex: 3 });
    const banners = await excelToPrisma.getData();
    console.log(banners);
    expect(banners.length).toBe(5);
  });

  it('should read parent table', async () => {
    await excelToPrisma.readSheet({ name: "user", rowNameIndex: 2, startRowIndex: 3 });
    const users = await excelToPrisma.getData();
    expect(users.length).toBe(5);
  });

  it('should linking one to many relationships', async () => {
    const findKeyArr = [2,3,5];
    await excelToPrisma.readSheet({ name: "user", rowNameIndex: 2, startRowIndex: 3 }).then( async (sheetOption) => {
      await excelToPrisma.oneToManyCreate({ name: "post", fk: 'userId', many: sheetOption.name, rowNameIndex: 2, startRowIndex: 3 });
    });
    const userPosts = await excelToPrisma.getData();
    
    expect(userPosts.filter(userPost => findKeyArr.includes(userPost.userId))).toContainEqual(
      expect.objectContaining({ post: { create: expect.any(Array) } })
    );
  });

  it('should linking subtables in a one to many relationship', async () => {
    const findKeyArr = [1,3,4,5];
    await excelToPrisma.readSheet({ name: "user", rowNameIndex: 2, startRowIndex: 3 }).then( async (sheetOption) => {
      await excelToPrisma.oneToManyCreate({ name: "product", fk: 'userId', many: sheetOption.name, rowNameIndex: 2, startRowIndex: 3, oneToOneOrManyOptions: [
        { key: 'infoProductTag', option: 'many', operation: 'connect'},
        { key: 'infoProductTag', option: 'one', operation: 'set'}
      ] }).then( async (sheetOption) => {
        await excelToPrisma.oneToManyCreate({ name: "productComment", fk: 'productId', many: sheetOption.name, rowNameIndex: 2, startRowIndex: 3 });
      });
    });
    const userProductComments = await excelToPrisma.getData();
    
    expect(userProductComments.filter(userProductComment => findKeyArr.includes(userProductComment.userId))).toContainEqual(
      expect.objectContaining({ 
        product: expect.objectContaining({ 
          create: expect.arrayContaining([
            expect.objectContaining({ 
              productComment: expect.objectContaining({ create: expect.any(Array) }) 
            })
          ]) 
        }) 
      })
    );
  });

  it('should linking subtables in a one to many in many relationship', async () => {
    const findKeyArr = [1,4,5];
    await excelToPrisma.readSheet({ name: "user", rowNameIndex: 2, startRowIndex: 3 }).then(async (sheetOption) => {
      await excelToPrisma.oneToManyCreate({ name: "product", fk: 'userId', many: sheetOption.name, rowNameIndex: 2, startRowIndex: 3, oneToOneOrManyOptions: [
        { key: 'infoProductTag', option: 'many', operation: 'connect'},
        { key: 'infoProductTag', option: 'one', operation: 'set'}
      ] }).then(async (sheetOption) => {
        await excelToPrisma.oneToManyCreate({ name: "productComment", fk: 'productId', many: sheetOption.name, rowNameIndex: 2, startRowIndex: 3 }).then(async (sheetOption) => {
          await excelToPrisma.oneToManyCreate({ name: "productCommentHistory", fk: "productCommentId", many: sheetOption.name, rowNameIndex: 2, startRowIndex: 3 });
        });
      });
    });
    const userProductCommentHistories = await excelToPrisma.getData();
  
    expect(userProductCommentHistories.filter(userProductCommentHistory => findKeyArr.includes(userProductCommentHistory.userId))).toContainEqual(
      expect.objectContaining({
        product: expect.objectContaining({
          create: expect.arrayContaining([
            expect.objectContaining({
              productComment: expect.objectContaining({
                create: expect.arrayContaining([
                  expect.objectContaining({
                    productCommentHistory: expect.objectContaining({ create: expect.any(Array) })
                  })
                ])
              })
            })
          ])
        })
      })
    );
  });
});