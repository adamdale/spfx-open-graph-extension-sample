import * as React from 'react';
import styles from './GraphExtensionsCrud.module.scss';
import { IGraphExtensionsCrudProps } from './IGraphExtensionsCrudProps';
import {
  DefaultButton,
  Stack,
  IStackTokens,
  MessageBar,
  MessageBarType,
  Text,
  Separator,
  Dropdown,
  IDropdownStyles,
  IDropdownOption
} from 'office-ui-fabric-react';

import useMsGraphProvider, { IGraphServices } from '../../../services/GraphService';

const stackTokens: IStackTokens = { childrenGap: 40 };

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};

const options: IDropdownOption[] = [
  { key: 'Concorde', text: 'Concorde' },
  { key: 'Jumbo Jet', text: 'Jumbo Jet' },
  { key: 'Bi Plane', text: 'Bi Plane' }
];

export const GraphExtensionsCrud: React.FunctionComponent<IGraphExtensionsCrudProps> = (
  props: IGraphExtensionsCrudProps
) => {
  const [favoritePlane, setFavoritePlane] = React.useState<string>("");
  const [createBtnDisabled, setCreateBtnDisabled] = React.useState<boolean>(true);
  const [readBtnDisabled, setReadBtnDisabled] = React.useState<boolean>(false);
  const [updateBtnDisabled, setUpdateBtnDisabled] = React.useState<boolean>(true);
  const [deleteBtnDisabled, setDeleteBtnDisabled] = React.useState<boolean>(true);
  const [messageBar, setMessageBar] = React.useState<MessageBarType>(MessageBarType.info);
  const [messageBarMessage, setMessageBarMessage] = React.useState<string>(props.description);
  const [msGraphProvider, setMSGraphProvider] = React.useState<IGraphServices>();

  const fetchMsGraphProvider = async () => {
    setMSGraphProvider(await useMsGraphProvider(props.context.msGraphClientFactory));
  };

  const _onChangeSelection = (ev: React.FormEvent<HTMLInputElement>, option?: any) => {
    setFavoritePlane(option.text);
    setMessageBarMessage(`Set Favorite Plane to ${option.text}?`);
    setMessageBar(MessageBarType.warning);
  }

  const _create = async () => {
    await msGraphProvider._createSchemaExtension("com.bestranet.roamingSettings", favoritePlane).then(() => {
      setCreateBtnDisabled(true);
      setReadBtnDisabled(false);
      setUpdateBtnDisabled(false);
      setMessageBarMessage(`Your favorite plane has been updated to ${favoritePlane}!`);
      setMessageBar(MessageBarType.success);
    });
  }

  const _read = async () => {
    await msGraphProvider._readSchemaExtension("com.bestranet.roamingSettings").then((value: any) => {
      if (!value) {
        setCreateBtnDisabled(false);
        setReadBtnDisabled(true);
        setUpdateBtnDisabled(true);
        setDeleteBtnDisabled(true);
        setFavoritePlane(`<Please Choose>`);
        setMessageBarMessage(`Please select a favorite plane!`);
        setMessageBar(MessageBarType.warning);
      }
      else {
        setUpdateBtnDisabled(false);
        setDeleteBtnDisabled(false);
        setFavoritePlane(value.plane ? value.plane : null);
        setMessageBarMessage(props.description);
        setMessageBar(MessageBarType.info);
      }
    });
  }

  const _update = async () => {
    await msGraphProvider._updateSchemaExtension("com.bestranet.roamingSettings", favoritePlane).then(() => {
      setMessageBarMessage(`Your favorite plane has been updated to ${favoritePlane}!`);
      setMessageBar(MessageBarType.success);
    });
  }

  const _delete = async () => {
    await msGraphProvider._deleteSchemaExtension("com.bestranet.roamingSettings").then(() => {
      setCreateBtnDisabled(false);
      setReadBtnDisabled(true);
      setUpdateBtnDisabled(true);
      setDeleteBtnDisabled(true);
      setMessageBarMessage("The favorite plane property has been removed");
      setMessageBar(MessageBarType.error);
    });
  }

  React.useEffect(() => {
    fetchMsGraphProvider();
  }, []);

  return (
    <div className={styles.graphExtensionsCrud}>
      <div className={styles.container}>
        <MessageBar messageBarType={messageBar} isMultiline={false} className={styles.notification}>
          {messageBarMessage}
        </MessageBar>
        <div>
          <Stack tokens={stackTokens}>
            <Text variant={'xxLarge'} block>My Favorite Plane is {favoritePlane}</Text>
          </Stack>
          <Separator />
          <Stack tokens={stackTokens}>
            <Dropdown
              placeholder="Select favorite"
              label="What's your favorite plane"
              options={options}
              styles={dropdownStyles}
              onChange={_onChangeSelection}
            />
          </Stack>
          <Separator />
          <Stack horizontal tokens={stackTokens} className={styles.stack}>
            <DefaultButton text="Create Favorite" className={styles.btn} onClick={_create} disabled={createBtnDisabled} />
            <DefaultButton text="Read Favorite" className={styles.btn} onClick={_read} disabled={readBtnDisabled} />
          </Stack>
          <Stack horizontal tokens={stackTokens}>
            <DefaultButton text="Update Favorite" className={styles.btn} onClick={_update} disabled={updateBtnDisabled} />
            <DefaultButton text="Delete Favorite" className={styles.btn} onClick={_delete} disabled={deleteBtnDisabled} />
          </Stack>
        </div>
      </div>
    </div>
  );

}