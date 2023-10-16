import ItemGaleryDto from './item-galery.dto';
import ItemCharacteristicDto from './item-characteristic.dto';

export default interface ItemDto {
  dp_name: string;
  dp_model: string;
  dp_cost: number;
  dp_photoUrl: string;
  dp_seoKeywords: string;
  dp_seoDescription: string;
  dp_itemCategoryId: number;
  dp_itemCharacteristics: ItemCharacteristicDto[];
  dp_itemGalery: ItemGaleryDto[];
}
