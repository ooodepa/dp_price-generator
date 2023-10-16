import * as ExcelJS from 'exceljs';

import env from './env';
import FetchItems from './utils/rest/api/v1/items/FetchItems';
import GetItemDto from './utils/rest/api/v1/items/dto/get-item.dto';
import FetchItemBrands from './utils/rest/api/v1/item-brands/FetchItemBrands';
import FetchItemCategories from './utils/rest/api/v1/item-categories/FetchItemCategories';

async function main() {
  let items: GetItemDto[] = [];
  for (let i = 0; i < env.backend__brandUrls.length; i++) {
    const jItems = (
      await FetchItems.get({ brand: env.backend__brandUrls[i] })
    ).sort((a, b) => a.dp_model.localeCompare(b.dp_model));
    items = [...items, ...jItems];
  }

  const brands = (await FetchItemBrands.get()).sort(
    (a, b) => a.dp_sortingIndex - b.dp_sortingIndex,
  );
  const categories = (await FetchItemCategories.get()).sort(
    (a, b) => a.dp_sortingIndex - b.dp_sortingIndex,
  );

  let rowId = 0;

  const itemRowIds: number[] = [];
  const categoryRowIds: number[] = [];
  const brandRowIds: number[] = [];

  const excelArray: string[][] = [];

  excelArray.push([
    'Картинка',
    'Модель',
    'Стоимость в Стамбуле',
    'Стоимость с НДС в Беларуси с доставкой',
    'Наименование',
  ]);
  rowId += 1;

  for (let i = 0; i < env.backend__brandUrls.length; ++i) {
    for (let j = 0; j < brands.length; ++j) {
      const jItBr = brands[j];
      if (env.backend__brandUrls[i] === jItBr.dp_urlSegment) {
        excelArray.push([jItBr.dp_name]);
        rowId += 1;
        brandRowIds.push(rowId);

        for (let k = 0; k < categories.length; ++k) {
          const jItCt = categories[k];
          if (brands[j].dp_id === jItCt.dp_itemBrandId && !jItCt.dp_isHidden) {
            excelArray.push([jItCt.dp_name]);
            rowId += 1;
            categoryRowIds.push(rowId);

            for (let q = 0; q < items.length; ++q) {
              const jIt = items[q];
              if (categories[k].dp_id === jIt.dp_itemCategoryId) {
                // img
                const imgs: string[] = jIt.dp_itemGalery.map(
                  e => e.dp_photoUrl,
                );
                imgs.unshift(jIt.dp_photoUrl);

                let image = '';
                const imageExtensions = /\.png$|\.jpg$|\.jpeg$/;
                for (let w = 0; w < imgs.length; ++w) {
                  if (imageExtensions.test(imgs[w])) {
                    image = imgs[w];
                    break;
                  }
                }

                const imgFormula =
                  image.length === 0 ? 'нет картинки' : `=IMAGE("${image}")`;
                // end img

                // cost
                const bynCost =
                  jIt.dp_cost === 0
                    ? 'уточняйте'
                    : Number(jIt.dp_cost).toFixed(2);

                let usdCostStr = '';
                for (let r = 0; r < jIt.dp_itemCharacteristics.length; ++r) {
                  const character = jIt.dp_itemCharacteristics[r];
                  if (character.dp_characteristicId === 24) {
                    usdCostStr = character.dp_value;
                    break;
                  }
                }
                const usdCost =
                  usdCostStr === ''
                    ? jIt.dp_cost > 0
                      ? ''
                      : 'уточняйте'
                    : Number(usdCostStr).toFixed(2);
                // end cost

                // name
                let nameEn = '';
                let nameRu = '';
                let nameTr = '';

                for (let r = 0; r < jIt.dp_itemCharacteristics.length; ++r) {
                  const character = jIt.dp_itemCharacteristics[r];
                  if (character.dp_characteristicId === 19) {
                    nameEn = character.dp_value;
                    break;
                  }
                }

                for (let r = 0; r < jIt.dp_itemCharacteristics.length; ++r) {
                  const character = jIt.dp_itemCharacteristics[r];
                  if (character.dp_characteristicId === 20) {
                    nameRu = character.dp_value;
                    break;
                  }
                }

                for (let r = 0; r < jIt.dp_itemCharacteristics.length; ++r) {
                  const character = jIt.dp_itemCharacteristics[r];
                  if (character.dp_characteristicId === 18) {
                    nameTr = character.dp_value;
                    break;
                  }
                }

                let name = '';

                if (nameTr.length > 0) {
                  name += `TR: ${nameTr}\n`;
                }

                if (nameEn.length > 0) {
                  name += `EN: ${nameEn}\n`;
                }

                if (nameRu.length > 0) {
                  name += `RU: ${nameRu}\n`;
                }

                if (name.length === 0) {
                  name = jIt.dp_name;
                }

                name = name.trim();
                // end name

                excelArray.push([
                  // Cols ABCD
                  imgFormula,
                  jIt.dp_model,
                  usdCost,
                  bynCost,
                  name,
                ]);
                rowId += 1;
                itemRowIds.push(rowId);
              }
            }
          }
        }
        break;
      }
    }
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('List1');

  // Заполнение ячеек данными
  worksheet.addRows(excelArray);

  for (let i = 0; i < excelArray.length; ++i) {
    const id = i + 1;

    // Отрисовываем границы яцеек
    ['A', 'B', 'C', 'D', 'E'].forEach(e => {
      worksheet.getCell(`${e}${id}`).border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
      };
      worksheet.getCell(`${e}${id}`).font = {
        size: 8,
      };
    });

    // переносить текст, если не вмешается в яцейку по ширине
    worksheet.getCell(`E${id}`).alignment = {
      wrapText: true,
      // vertical: 'top',
      // horizontal: 'left',
    };
  }

  // Выравнивание по центру (картинку)
  for (let i = 0; i < itemRowIds.length; ++i) {
    const id = itemRowIds[i];
    if (!id) continue;

    worksheet.getCell(`A${id}`).alignment = {
      horizontal: 'center',
      vertical: 'middle',
    };
  }

  // Выравнивание (модели)
  for (let i = 0; i < itemRowIds.length; ++i) {
    const id = itemRowIds[i];
    if (!id) continue;

    worksheet.getCell(`B${id}`).alignment = {
      horizontal: 'left',
      vertical: 'middle',
    };
  }

  // Выравнивание по правому краю (цены)
  for (let i = 0; i < itemRowIds.length; ++i) {
    const id = itemRowIds[i];
    if (!id) continue;
    ['C', 'D'].forEach(e => {
      worksheet.getCell(`${e}${id}`).alignment = {
        horizontal: 'right',
        vertical: 'middle',
      };
    });
  }

  // Выравнивание (Наименования)
  for (let i = 0; i < itemRowIds.length; ++i) {
    const id = itemRowIds[i];
    if (!id) continue;

    worksheet.getCell(`E${id}`).alignment = {
      wrapText: true,
      horizontal: 'left',
      vertical: 'middle',
    };
  }

  // Объединение ячеек (для наименования бренда)
  for (let i = 0; i <= brandRowIds.length; i++) {
    const id = brandRowIds[i];
    if (!id) continue;

    worksheet.mergeCells(`A${id}:E${id}`);

    worksheet.getCell(`A${id}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'b6d7a8' },
    };

    worksheet.getCell(`A${id}`).alignment = {
      wrapText: true,
      horizontal: 'center',
      vertical: 'middle',
    };
  }

  // Объединение ячеек (для наименования категории)
  for (let i = 0; i <= categoryRowIds.length; i++) {
    const id = categoryRowIds[i];
    if (!id) continue;

    worksheet.mergeCells(`A${id}:E${id}`);

    worksheet.getCell(`A${id}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'd9ead3' },
    };

    worksheet.getCell(`A${id}`).alignment = {
      wrapText: true,
      horizontal: 'center',
      vertical: 'middle',
    };
  }

  // Установка ширины столбцов
  worksheet.getColumn('A').width = 20;
  worksheet.getColumn('B').width = 10;
  worksheet.getColumn('C').width = 10;
  worksheet.getColumn('D').width = 10;
  worksheet.getColumn('E').width = 90;

  // Установка ширины столбцов
  for (let i = 1; i < excelArray.length; ++i) {
    worksheet.getRow(i + 1).height = 40;
  }

  // Стили для заголовка
  ['A', 'B', 'C', 'D', 'E'].forEach(e => {
    worksheet.getCell(`${e}1`).alignment = {
      wrapText: true,
      horizontal: 'center',
      vertical: 'middle',
    };

    worksheet.getCell(`${e}1`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'f9cb9c' },
    };
  });

  worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

  await workbook.xlsx.writeFile('result.xlsx');
}

main();
