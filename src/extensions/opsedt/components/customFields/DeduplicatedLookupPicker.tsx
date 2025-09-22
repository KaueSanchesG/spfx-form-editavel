import * as React from 'react';
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IDynamicFieldProps } from '@pnp/spfx-controls-react/lib/controls/dynamicForm/dynamicField';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ITag, TagPicker } from '@fluentui/react/lib/Pickers';

import styles from '../Opsedt.module.scss';

export interface IDeduplicatedLookupPickerProps {
    fieldProps: IDynamicFieldProps;
    sp: SPFI;
    listGuid: string;
    lookupField: string;
    onSelectionChange: (ids: number[]) => void;
}

export const DeduplicatedLookupPicker: React.FC<IDeduplicatedLookupPickerProps> = (props) => {
    const { fieldProps, sp, listGuid, lookupField, onSelectionChange } = props;

    const [allItemsAsTags, setAllItemsAsTags] = React.useState<ITag[]>([]);
    const [isLoading, setIsLoading] = React.useState<boolean>(true);

    const initialValues = React.useMemo((): ITag[] => {
        if (!fieldProps.value || !Array.isArray(fieldProps.value)) { return []; }
        return fieldProps.value
            .filter(val => val && typeof val.key !== 'undefined' && val.key !== null)
            .map((val: any) => ({ key: val.key.toString(), name: val.name }));
    }, [fieldProps.value]);

    const [selectedItems, setSelectedItems] = React.useState<ITag[]>(initialValues);


    React.useEffect(() => {
        const fetchItems = async () => {
            setIsLoading(true);
            try {
                const items: any[] = await sp.web.lists.getById(listGuid).items();

                const filesOnly = items.filter(item => item.FileSystemObjectType === 0);

                const uniqueMap = new Map<string, any>();
                for (const item of filesOnly) {
                    const key = item[lookupField];

                    if (key && !uniqueMap.has(key)) {
                        uniqueMap.set(key, item);
                    }
                }
                const uniqueItems = Array.from(uniqueMap.values());

                const tagOptions: ITag[] = uniqueItems.map(item => ({
                    key: item.ID.toString(),
                    name: item[lookupField]
                }));
                setAllItemsAsTags(tagOptions);

            } catch (error) {
                console.error(`ERRO CRÍTICO ao buscar dados para o campo '${props.fieldProps.label}':`, error);
            } finally {
                setIsLoading(false);
            }
        };

        fetchItems();
    }, [sp, listGuid, lookupField, props.fieldProps.label]);

    const onFilterChanged = (filterText: string, currentSelectedItems?: ITag[]): ITag[] => {
        const localSelectedItems = currentSelectedItems || [];
        return allItemsAsTags.filter(
            tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) !== -1 &&
                !localSelectedItems.some(item => item.key === tag.key)
        );
    };

    const onEmptyResolveSuggestions = (currentSelectedItems?: ITag[]): ITag[] => {
        const localSelectedItems = currentSelectedItems || [];
        return allItemsAsTags.filter(
            tag => !localSelectedItems.some(item => item.key === tag.key)
        );
    };

    const getTextFromItem = (item: ITag): string => item.name;

    const onChange = (items?: ITag[]) => {
        const newItems = items || [];
        setSelectedItems(newItems);
        const numericValues = newItems.map(item => parseInt(item.key.toString(), 10));
        if (onSelectionChange) {
            onSelectionChange(numericValues);
        }
    };

    if (isLoading) {
        return <Spinner size={SpinnerSize.small} label={`Carregando ${props.fieldProps.label}...`} />;
    }

    return (
        <div>
            <label className={styles.fieldLabel}>{props.fieldProps.label}</label>
            <TagPicker
                onResolveSuggestions={onFilterChanged}
                onEmptyResolveSuggestions={onEmptyResolveSuggestions}
                getTextFromItem={getTextFromItem}
                selectedItems={selectedItems}
                onChange={onChange}
                disabled={props.fieldProps.disabled}
            />
        </div>
    );
};