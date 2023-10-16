import ItemCharacteristicDto from './item-characteristic.dto';
import ItemGaleryDto from './item-galery.dto';
import ItemDto from './item.dto';

interface ItemCharacteristicWithIdDto extends ItemCharacteristicDto {
  dp_id: number;
}

interface ItemGaleryWithIdDto extends ItemGaleryDto {
  dp_id: number;
}

export default interface ItemWithIdDto extends ItemDto {
  dp_itemCharacteristics: ItemCharacteristicWithIdDto[];
  dp_itemGalery: ItemGaleryWithIdDto[];
}
