import * as React from 'react';
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/security";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import { PermissionKind } from '@pnp/sp/security';
import { Panel, PanelType, IPanelProps } from '@fluentui/react/lib/Panel';
import { IconButton } from '@fluentui/react/lib/Button';
import { DynamicForm } from '@pnp/spfx-controls-react/lib/controls/dynamicForm';

import styles from './Opsedt.module.scss';
import { formTitleResources } from './customFields/textResources';
import { FilteredMultiLookupPicker, IFilteredMultiLookupPickerRef } from './customFields/FilteredMultiLookupPicker';
import { CustomLookupEditor } from './customFields/CustomLookupEditor';
import { DeduplicatedLookupPicker } from './customFields/DeduplicatedLookupPicker';
import { IDynamicFieldProps, DynamicField } from '@pnp/spfx-controls-react/lib/controls/dynamicForm/dynamicField';

export interface IOpsedtProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

export interface IOpsedtState {
  isEditMode: boolean;
  canEdit: boolean;
  isSaving?: boolean;
  areaGestoraValue?: string;
}

export default class Opsedt extends React.Component<IOpsedtProps, IOpsedtState> {
  private sp: SPFI;
  private customLookupValue: number | null = null;
  private arquivosRelacionadosValue: number[] = [];
  private arquivosRelacionadosPickerRef = React.createRef<IFilteredMultiLookupPickerRef>();
  private normativosCanceladosValue: number[] = [];
  private normativoComplementarValue: number[] = [];

  constructor(props: IOpsedtProps) {
    super(props);
    this.sp = spfi(this.props.context.pageContext.web.absoluteUrl).using(SPFx(this.props.context));

    this.state = {
      isEditMode: this.props.displayMode === FormDisplayMode.Edit || this.props.displayMode === FormDisplayMode.New,
      canEdit: false,
      isSaving: false,
      areaGestoraValue: undefined,
    };
  }

  public async componentDidMount(): Promise<void> {
    if (this.props.displayMode === FormDisplayMode.Edit || this.props.displayMode === FormDisplayMode.New) {
      this.setState({ canEdit: true });
      return;
    }

    if (this.props.displayMode === FormDisplayMode.Display && this.props.context.itemId) {
      try {
        const list = this.sp.web.lists.getById(this.props.context.list.guid.toString());
        const item = list.items.getById(this.props.context.itemId);
        const perms = await item.getCurrentUserEffectivePermissions();
        const hasEditPermission = this.sp.web.hasPermissions(perms, PermissionKind.EditListItems);

        this.setState({ canEdit: hasEditPermission });

      } catch (error) {
        console.error("Erro ao verificar permissões no item:", error);
      }
    }
  }

  private handleLookupChange = (id: number | null): void => {
    this.customLookupValue = id;
  }

  private handleArquivosRelacionadosChange = (ids: number[]): void => {
    this.arquivosRelacionadosValue = ids;
  }

  private handleNormativosCanceladosChange = (ids: number[]): void => {
    this.normativosCanceladosValue = ids;
  }

  private handleNormativoComplementarChange = (ids: number[]): void => {
    this.normativoComplementarValue = ids;
  }

  private _onSubmitted = async (formData: any): Promise<void> => {
    this.setState({ isSaving: true });

    let newFileIds: number[] = [];
    if (this.arquivosRelacionadosPickerRef.current) {
      try {
        newFileIds = await this.arquivosRelacionadosPickerRef.current.uploadFiles();
      } catch (error) {
        console.error("Falha no processo de upload. O salvamento foi abortado.", error);
        this.setState({ isSaving: false });
        return;
      }
    }

    const allIds = this.arquivosRelacionadosValue.concat(newFileIds);
    const combinedIds = allIds.filter((id, index) => allIds.indexOf(id) === index);

    const finalPayload = {
      ...formData,
      aplicacaoNormativoId: this.customLookupValue,
      ArquivosRelacionadosId: combinedIds,
      NormativosCanceladosId: this.normativosCanceladosValue,
      normativo_x002d_complementarId: this.normativoComplementarValue
    };

    delete finalPayload.ArquivosRelacionados;
    delete finalPayload.NormativosCancelados;
    delete finalPayload.normativo_x002d_complementar;

    console.log("Payload FINAL (corrigido) enviado para o SharePoint:", finalPayload);

    try {
      if (this.props.displayMode === FormDisplayMode.New) {
        await this.sp.web.lists.getById(this.props.context.list.guid.toString()).items.add(finalPayload);
      } else if (this.props.context.itemId) {
        await this.sp.web.lists.getById(this.props.context.list.guid.toString())
          .items.getById(this.props.context.itemId)
          .update(finalPayload);
      }

      if (this.props.displayMode === FormDisplayMode.Display) {
        this.setState({ isEditMode: false, isSaving: false });
      } else {
        this.props.onSave();
      }

    } catch (error) {
      console.error("Erro ao salvar o item:", error);
      this.setState({ isSaving: false });
    }
  }

  private _switchToEditMode = (): void => {
    this.setState({ isEditMode: true });
  }

  private _onRenderNavigationContent = (props: IPanelProps, defaultRender?: (props?: IPanelProps) => JSX.Element | null): JSX.Element | null => {
    return (
      <div style={{ display: 'flex', alignItems: 'center', flexGrow: 1 }}>
        {this.props.displayMode === FormDisplayMode.Display && this.state.canEdit && !this.state.isEditMode && (
          <IconButton
            iconProps={{ iconName: 'Edit' }}
            title="Editar"
            ariaLabel="Editar"
            onClick={this._switchToEditMode}
          />
        )}
        <div style={{ flexGrow: 1 }}></div>
        {defaultRender && defaultRender(props)}
      </div>
    );
  }

  public render(): React.ReactElement<IOpsedtProps> {
    const { context, displayMode, onClose } = this.props;
    const currentDisplayMode = this.state.isEditMode ? FormDisplayMode.Edit : displayMode;

    const lang = navigator.language.toLowerCase().startsWith('es') ? 'ES' : 'PT';
    const modeText = currentDisplayMode === FormDisplayMode.New ? formTitleResources.new[lang] : currentDisplayMode === FormDisplayMode.Edit ? formTitleResources.edit[lang] : formTitleResources.view[lang];

    const fieldsToOrder: string[] = ["Identificador", "TituloPT", "TituloES1", "Area_x0020_Gestora", "siglaDoTipoDoNormativo", "aplicacaoNormativo", "DataInicioVigencia0", "DataFimVigencia", "normativo_x002d_complementar", "normativoEmAnexo_historico", "ArquivosRelacionados", "NormativosCancelados", "historicocancelado", "Revisor", "Gestor", "MotivoRevisao", "Parecer", "comentario"];
    const fieldsToHide: string[] = ["FileLeafRef", "Numeral", "Correcao", "AreaGestoraTexto", "tramitacao", "DocumentosEmTramitacao", "HistoricoNormativo", "Revisao", "IdiomaNormativo", "IdiomaVersao", "NotasDaRevisao", "motivoCancelamento", "EstadoNormativo", "DescricaoTipoDocumento", "corretores"];
    const fieldsToDisable: string[] = ["FileLeafRef", "Numeral", "Identificador", "Area_x0020_Gestora", "siglaDoTipoDoNormativo", "comentario", "normativoEmAnexo_historico", "Revisao", "comentario", "historicocancelado", "Parecer"];

    const fieldOverrides = {
      Area_x0020_Gestora: (fieldProps: IDynamicFieldProps) => {
        return (
          <DynamicField
            {...fieldProps}
            onChanged={(value: any) => {
              this.setState({ areaGestoraValue: value });
              if (fieldProps.onChanged) {
                fieldProps.onChanged(fieldProps.columnInternalName, value, true);
              }
            }}
          />
        );
      },
      aplicacaoNormativo: (fieldProps: IDynamicFieldProps) => {
        return (
          <CustomLookupEditor
            sp={this.sp}
            fieldProps={fieldProps}
            listGuid="951911ca-bb5a-4de6-b166-4d22431eb6af"
            onSelectionChange={this.handleLookupChange}
          />
        );
      },
      ArquivosRelacionados: (fieldProps: IDynamicFieldProps) => {
        return (
          <FilteredMultiLookupPicker
            ref={this.arquivosRelacionadosPickerRef}
            context={this.props.context}
            fieldProps={fieldProps}
            sp={this.sp}
            fieldInternalName="ArquivosRelacionados"
            onSelectionChange={this.handleArquivosRelacionadosChange}
            areaGestora={this.state.areaGestoraValue}
          />
        );
      },
      NormativosCancelados: (fieldProps: IDynamicFieldProps) => {
        return (
          <DeduplicatedLookupPicker
            fieldProps={fieldProps}
            sp={this.sp}
            listGuid={"2d235dcd-cb76-46aa-833b-4f99005e6f17"}

            lookupField="Identificador"
            onSelectionChange={this.handleNormativosCanceladosChange}
          />
        );
      },

      normativo_x002d_complementar: (fieldProps: IDynamicFieldProps) => {
        return (
          <DeduplicatedLookupPicker
            fieldProps={fieldProps}
            sp={this.sp}
            listGuid={"2d235dcd-cb76-46aa-833b-4f99005e6f17"}
            lookupField="Identificador"
            onSelectionChange={this.handleNormativoComplementarChange}
          />
        );
      }
    };

    return (
      <Panel
        isOpen={true}
        type={PanelType.large}
        headerText={`${modeText}`}
        onDismiss={onClose}
        isLightDismiss={true}
        onRenderNavigationContent={this._onRenderNavigationContent}
        isFooterAtBottom={true}
      >
        <div className={styles.normativosVigentes}>
          <DynamicForm
            key={this.state.isEditMode ? 'edit-form' : 'display-form'}
            context={context as any}
            listId={context.list.guid.toString()}
            listItemId={displayMode === FormDisplayMode.New ? undefined : context.itemId}
            onSubmitted={(formData) => this._onSubmitted(formData)}
            onCancelled={() => {
              if (this.props.displayMode === FormDisplayMode.Display && this.state.isEditMode) {
                this.setState({ isEditMode: false });
              } else {
                this.props.onClose();
              }
            }}
            fieldOverrides={fieldOverrides as any}
            fieldOrder={fieldsToOrder}
            hiddenFields={fieldsToHide}
            disabledFields={fieldsToDisable}
            disabled={!this.state.isEditMode || this.state.isSaving}
          />
        </div>
      </Panel>
    );
  }
}