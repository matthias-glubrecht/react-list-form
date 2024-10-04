import * as React from 'react';
import { ISPFormFieldProps } from './SPFormField';
//import { PeoplePicker, PrincipalType } from 'office-ui-fabric-react/lib/PeoplePicker';
import * as strings from 'FormFieldStrings';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';

const SPFieldUserEdit: React.SFC<ISPFormFieldProps> = (props) => {
    // We need to set value to empty string when null or undefined to force TextField still be used like a controlled component
    let value;
    if (props.value)
    {
        if (typeof props.value === "string") {
            const o = JSON.parse(props.value);
            if (o[0].key) {
                value = o[0].key;
            }
            else {
                value = o;
            }
        }
        else if (typeof props.value === "object") {
            value = props.value[0].Key;
        }
    }
    else {
        value = '';
    }

    return <PeoplePicker
        //peoplePickerWPclassName={styles.dropdown}
        context={props.context}
        personSelectionLimit={1}
        showtooltip={true}
        principalTypes={[PrincipalType.User]}
        resolveDelay={1000}
        defaultSelectedUsers={[value]}
        selectedItems={(pickerProps) => {
            if (pickerProps.length > 0) {
                const value = JSON.stringify([{key: pickerProps[0].id, IsResolved: true}]);
                props.valueChanged(value);
            }
            else {
                props.valueChanged('');
            }
        }}
    />
};

export default SPFieldUserEdit;
