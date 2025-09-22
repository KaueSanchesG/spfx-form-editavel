import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IDynamicFieldProps } from '@pnp/spfx-controls-react/lib/controls/dynamicForm/dynamicField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { textResources } from './textResources';

import styles from '../Opsedt.module.scss';

export interface ICustomLookupEditorProps {
    sp: SPFI;
    fieldProps: IDynamicFieldProps;
    listGuid: string;
    onSelectionChange: (id: number | null) => void;
}

export const CustomLookupEditor: React.FC<ICustomLookupEditorProps> = (props) => {
    const { sp, fieldProps, listGuid, onSelectionChange } = props;
    const { disabled, value, required } = fieldProps;

    const [loading, setLoading] = useState(false);
    const [options, setOptions] = useState<IDropdownOption[]>([]);
    const [selectedKey, setSelectedKey] = useState<number | null | undefined>();
    const [isInitialized, setIsInitialized] = useState(false);

    const lang = navigator.language.toLowerCase().startsWith('es') ? 'ES' : 'PT';

    useEffect(() => {
        if (value && !isInitialized) {
            const lookupValue = Array.isArray(value) ? value[0] : value;
            const initialId = lookupValue?.lookupId ?? lookupValue?.key;
            if (initialId) {
                setSelectedKey(initialId);
                onSelectionChange(initialId);
                setIsInitialized(true);
            }
        }
    }, [value, isInitialized, onSelectionChange]);

    useEffect(() => {
        setLoading(true);
        const fetchOptions = async () => {
            try {
                const items: any[] = await sp.web.lists.getById(listGuid).items.select("ID", "DescTipoAplicacaoES", "DescTipoAplicacaoPT")();
                const dropdownOptions = items.map(i => ({
                    key: i.ID,
                    text: `${String(i.ID).padStart(2, '0')} - ${lang === 'ES' ? (i.DescTipoAplicacaoES || i.DescTipoAplicacaoPT) : (i.DescTipoAplicacaoPT || i.DescTipoAplicacaoES)}`
                }));
                setOptions(dropdownOptions);
            } catch (error) {
                console.error(error);
            } finally {
                setLoading(false);
            }
        };
        fetchOptions();
    }, [sp, listGuid, lang]);

    const handleOnChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
        const newKey = option ? option.key as number : null;
        setSelectedKey(newKey);
        onSelectionChange(newKey);
    };


    return (
        <div>
            <label className={styles.fieldLabel}>{textResources.catLabel[lang]}</label>

            <Dropdown
                required={required}
                options={options}
                selectedKey={selectedKey}
                onChange={handleOnChange}
                disabled={disabled || loading}
                placeholder={loading ? textResources.loading[lang] : textResources.selectOption[lang]}
            />
        </div>
    );
};