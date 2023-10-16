import ItemCategoryDto from './item-category.dto';

export default interface ItemCategoryWithId extends ItemCategoryDto {
  dp_id: number;
}
