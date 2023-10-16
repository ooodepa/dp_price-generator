import axios from 'axios';

import GetItemDto from './dto/get-item.dto';

interface GetItemFilterDto {
  brand?: string;
}

export default class FetchItems {
  static async get(query: GetItemFilterDto = {}) {
    try {
      let url = 'https://de-pa.by/api/v1/items?';
      if (query.brand) {
        url += `brand=${query.brand}`;
      }

      const response = await axios.get(url);
      const json: GetItemDto[] = response.data;
      return json;
    } catch (exception) {
      return [];
    }
  }
}
