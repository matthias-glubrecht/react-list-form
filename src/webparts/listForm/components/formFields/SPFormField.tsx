import * as React from 'react';
import { ControlMode } from '../../../../common/datatypes/ControlMode';
import { IFieldSchema } from '../../../../common/services/datatypes/RenderListData';

import FormField from './FormField';
import { IFormFieldProps } from './FormField';
import { IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

import SPFieldTextEdit from './SPFieldTextEdit';
import SPFieldLookupEdit from './SPFieldLookupEdit';
import SPFieldChoiceEdit from './SPFieldChoiceEdit';
import SPFieldNumberEdit from './SPFieldNumberEdit';
import SPFieldDateEdit from './SPFieldDateEdit';
import SPFieldBooleanEdit from './SPFieldBooleanEdit';
import SPFieldTextDisplay from './SPFieldTextDisplay';
import SPFieldLookupDisplay from './SPFieldLookupDisplay';
import SPFieldUserDisplay from './SPFieldUserDisplay';
import SPFieldUrlDisplay from './SPFieldUrlDisplay';

import * as strings from 'FormFieldStrings';
import styles from './SPFormField.module.scss';
import SPFieldUserEdit from './SPFieldUserEdit';


const EditFieldTypeMappings: {[fieldType: string]: React.StatelessComponent<ISPFormFieldProps>} = {
  Text: SPFieldTextEdit,
  Note: SPFieldTextEdit,
  Lookup: SPFieldLookupEdit,
  LookupMulti: SPFieldLookupEdit,
  Choice: SPFieldChoiceEdit,
  MultiChoice: SPFieldChoiceEdit,
  Number: SPFieldNumberEdit,
  Currency: SPFieldNumberEdit,
  DateTime: SPFieldDateEdit,
  Boolean: SPFieldBooleanEdit,
  File: SPFieldTextEdit,
  User: SPFieldUserEdit,
  /* The following are known but unsupported types as of now:
  UserMulti: null,
  URL: null,
  TaxonomyFieldType: null,
  Attachments: null,
  TaxonomyFieldTypeMulti: null,
  */
};


const DisplayFieldTypeMappings: {[fieldType: string]: {component: React.StatelessComponent<ISPFormFieldProps>, valuePreProcess?: (value: any) => any}} = {
  Text: { component: SPFieldTextDisplay },
  Note: { component: SPFieldTextDisplay },
  Lookup: { component: SPFieldLookupDisplay },
  LookupMulti: { component: SPFieldLookupDisplay },
  Choice: { component: SPFieldTextDisplay },
  MultiChoice: {component: SPFieldTextDisplay, valuePreProcess: (val) => val ? val.join(', ') : '' },
  Number: { component: SPFieldTextDisplay },
  Currency: { component: SPFieldTextDisplay },
  DateTime: { component: SPFieldTextDisplay },
  Boolean: { component: SPFieldTextDisplay },
  User: { component: SPFieldUserDisplay },
  UserMulti: { component: SPFieldUserDisplay },
  URL: { component: SPFieldUrlDisplay },
  File: { component: SPFieldTextDisplay},
  TaxonomyFieldType: { component: SPFieldTextDisplay, valuePreProcess: (val) => val ? val.Label : '' },
  TaxonomyFieldTypeMulti: { component: SPFieldTextDisplay, valuePreProcess: (val) => val ? val.map( (v) => v.Label ).join(', ') : '' },
  /* The following are known but unsupported types as of now:
  Attachments: null,
  */
};

export interface ISPFormFieldProps extends IFormFieldProps {
    extraData?: any;
    fieldSchema: IFieldSchema;
    hideIfFieldUnsupported?: boolean;
}


const SPFormField: React.SFC<ISPFormFieldProps> = (props) => {
  let fieldControl = null;
  const fieldType = props.fieldSchema.FieldType;
  if (props.controlMode === ControlMode.Display) {
    if (DisplayFieldTypeMappings.hasOwnProperty(fieldType)) {
      const fieldMapping = DisplayFieldTypeMappings[fieldType];
      const childProps = fieldMapping.valuePreProcess ? {...props, value: fieldMapping.valuePreProcess(props.value)} : props;
      fieldControl = React.createElement( fieldMapping.component, childProps );
    } else if (!props.hideIfFieldUnsupported) {
      const value = (props.value) ? ((typeof props.value === 'string') ? props.value : JSON.stringify(props.value)) : '';
      fieldControl = <div className={`ard-${fieldType}field-display`}>
          <span>{value}</span>
          <div className={styles.unsupportedFieldMessage}><Icon iconName='Error' />{`${strings.UnsupportedFieldType} "${fieldType}"`}</div>
        </div>;
    }
  } else {
    if (EditFieldTypeMappings.hasOwnProperty(fieldType)) {
      fieldControl = React.createElement( EditFieldTypeMappings[fieldType], props );
    } else if (!props.hideIfFieldUnsupported) {
      const isObjValue = (props.value) && (typeof props.value !== 'string');
      const value = (props.value) ? ((typeof props.value === 'string') ? props.value : JSON.stringify(props.value)) : '';
      fieldControl = <TextField
                readOnly
                multiline={isObjValue}
                value={value}
                errorMessage={`${strings.UnsupportedFieldType} "${fieldType}"`}
                underlined
              />;
    }
  }
  return (fieldControl)
    ? <FormField
            {...props}
            label={props.label || props.fieldSchema.Title}
            description={props.description || props.fieldSchema.Description}
            required={props.fieldSchema.Required}
            errorMessage={props.errorMessage}
            >
            {fieldControl}
      </FormField>
    : null;
};


export default SPFormField;
