import axios from 'axios';

import GetItemCategoryDto from './dto/get-item-category.dto';

export default class FetchItemCategories {
  static async get() {
    try {
      let url = 'https://de-pa.by/api/v1/item-categories';
      const response = await axios.get(url);
      const json: GetItemCategoryDto[] = response.data;
      return json;
    } catch (exception) {
      return [];
    }
  }
}
