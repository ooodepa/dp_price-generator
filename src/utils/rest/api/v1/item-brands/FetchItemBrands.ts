import axios from 'axios';

import GetItemBrandDto from './dto/get-item-category.dto';

export default class FetchItemBrands {
  static async get() {
    try {
      let url = 'https://de-pa.by/api/v1/item-brands';
      const response = await axios.get(url);
      const json: GetItemBrandDto[] = response.data;
      return json;
    } catch (exception) {
      return [];
    }
  }
}
