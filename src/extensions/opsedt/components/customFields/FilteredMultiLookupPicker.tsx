import * as React from 'react';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { IDynamicFieldProps } from '@pnp/spfx-controls-react/lib/controls/dynamicForm/dynamicField';
import { ITag, TagPicker } from '@fluentui/react/lib/Pickers';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import { DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Icon } from '@fluentui/react/lib/Icon';

import { textResources } from './textResources';
import styles from '../Opsedt.module.scss';

export interface IFilteredMultiLookupPickerRef {
    uploadFiles: () => Promise<number[]>;
}

export interface IFilteredMultiLookupPickerProps {
    context: BaseComponentContext;
    fieldProps: IDynamicFieldProps;
    sp: SPFI;
    fieldInternalName: string;
    onSelectionChange: (ids: number[]) => void;
    areaGestora?: string;
}

export const FilteredMultiLookupPicker = React.forwardRef<IFilteredMultiLookupPickerRef, IFilteredMultiLookupPickerProps>((props, ref) => {

    const { fieldProps, sp, onSelectionChange } = props;
    const targetListId = "25b65cb4-d567-40ec-b440-2f0dff0af8c1";

    const lang = navigator.language.toLowerCase().startsWith('es') ? 'ES' : 'PT';

    const [allFileTags, setAllFileTags] = React.useState<ITag[]>([]);
    const [isLoading, setIsLoading] = React.useState<boolean>(true);
    const [filesToUpload, setFilesToUpload] = React.useState<File[]>([]);
    const fileInputRef = React.useRef<HTMLInputElement>(null);
    const initialValues = React.useMemo((): ITag[] => {
        if (!fieldProps.value || !Array.isArray(fieldProps.value)) {
            return [];
        }
        return fieldProps.value
            .filter(val => val && typeof val.key !== 'undefined' && val.key !== null)
            .map((val: any) => ({
                key: val.key.toString(),
                name: val.name
            }));
    }, [fieldProps.value]);
    const [selectedItems, setSelectedItems] = React.useState<ITag[]>(initialValues);

    React.useEffect(() => {
        sp.web.lists.getById(targetListId).items
            .select("Id", "FileLeafRef", "Title", "FileSystemObjectType")()
            .then((allItems: any[]) => {
                const filesOnly = allItems.filter(item => item.FileSystemObjectType === 0);
                const tagOptions: ITag[] = filesOnly.map(file => ({
                    key: file.Id.toString(),
                    name: file.Title || file.FileLeafRef
                }));
                setAllFileTags(tagOptions);
            })
            .catch((error: any) => {
                console.error("Erro ao carregar arquivos existentes:", error);
            })
            .finally(() => {
                setIsLoading(false);
            });
    }, []);

    React.useImperativeHandle(ref, () => ({
        uploadFiles: async (): Promise<number[]> => {
            if (filesToUpload.length === 0) {
                return [];
            }
            const uploadedFileIds: number[] = [];
            const targetLibrary = sp.web.lists.getById(targetListId);
            for (const file of filesToUpload) {
                try {
                    const fileInfoResult = await targetLibrary.rootFolder.files.addChunked(file.name, file, { Overwrite: true });
                    const fileObject = await sp.web.getFileByServerRelativePath(fileInfoResult.ServerRelativeUrl);
                    const itemToUpdate = await fileObject.getItem();
                    await itemToUpdate.update({ Title: file.name });
                    const itemData = await itemToUpdate.select("Id")();
                    uploadedFileIds.push(itemData.Id);
                } catch (error) {
                    console.error(`Erro ao fazer upload do arquivo ${file.name}:`, error);
                }
            }
            return uploadedFileIds;
        }
    }));

    const handleFileSelection = (event: React.ChangeEvent<HTMLInputElement>): void => {
        if (event.target.files) {
            const newFiles = Array.from(event.target.files);
            const uniqueNewFiles = newFiles.filter(newFile =>
                !filesToUpload.some(existingFile => existingFile.name === newFile.name)
            );
            setFilesToUpload(prevFiles => [...prevFiles, ...uniqueNewFiles]);
        }
    };

    const handleRemoveFile = (fileName: string): void => {
        setFilesToUpload(prevFiles => prevFiles.filter(file => file.name !== fileName));
    };

    const onFilterChanged = (filterText: string, currentSelectedItems?: ITag[]): ITag[] => {
        const localSelectedItems = currentSelectedItems || [];
        return allFileTags.filter(
            tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) !== -1 &&
                !localSelectedItems.some(item => item.key === tag.key)
        );
    };

    const onEmptyResolveSuggestions = (currentSelectedItems?: ITag[]): ITag[] => {
        const localSelectedItems = currentSelectedItems || [];
        return allFileTags.filter(
            tag => !localSelectedItems.some(item => item.key === tag.key)
        );
    };

    const onChange = (items?: ITag[]) => {
        const newItems = items || [];
        setSelectedItems(newItems);
        const numericValues = newItems.map(item => parseInt(item.key.toString(), 10));
        if (onSelectionChange) {
            onSelectionChange(numericValues);
        }
    };

    const getTextFromItem = (item: ITag): string => item.name;

    if (isLoading) {
        return (
            <div>
                <label>{fieldProps.label}</label>
                <Spinner size={SpinnerSize.small} label="Carregando arquivos..." />
            </div>
        );
    }

    return (
        <div>
            <label className={styles.fieldLabel}>{fieldProps.label}</label>
            <TagPicker
                onResolveSuggestions={onFilterChanged}
                onEmptyResolveSuggestions={onEmptyResolveSuggestions}
                getTextFromItem={getTextFromItem}
                selectedItems={selectedItems}
                onChange={onChange}
                disabled={fieldProps.disabled}
            />
            {!fieldProps.disabled && (
                <div style={{ marginTop: '15px' }}>
                    <input
                        type="file"
                        multiple
                        ref={fileInputRef}
                        style={{ display: 'none' }}
                        onChange={handleFileSelection}
                    />
                    <DefaultButton
                        iconProps={{ iconName: 'Upload' }}
                        text="Carregar Novos Arquivos"
                        onClick={() => fileInputRef.current?.click()}
                    />

                    {filesToUpload.length > 0 && (
                        <div style={{ marginTop: '10px' }}>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 5 }} style={{ marginBottom: '10px', color: '#a80000' }}>
                                <Icon iconName="Warning" />
                                <span style={{ fontSize: '12px' }}>
                                    {textResources.fileOverwriteWarning[lang]}
                                </span>
                            </Stack>

                            <strong>Arquivos para enviar:</strong>
                            <Stack tokens={{ childrenGap: 5 }} styles={{ root: { marginTop: 5 } }}>
                                {filesToUpload.map((file, index) => (
                                    <Stack horizontal key={index} verticalAlign="center">
                                        <span>{file.name}</span>
                                        <IconButton
                                            iconProps={{ iconName: 'Cancel' }}
                                            title="Remover"
                                            ariaLabel="Remover"
                                            onClick={() => handleRemoveFile(file.name)}
                                        />
                                    </Stack>
                                ))}
                            </Stack>
                        </div>
                    )}
                </div>
            )}
        </div>
    );
});