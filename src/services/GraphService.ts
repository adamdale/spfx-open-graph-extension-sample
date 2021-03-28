import { MSGraphClientFactory } from '@microsoft/sp-http';

export interface IGraphServices {
    _createSchemaExtension(extensionName: string, meTaskSettings: string): Promise<any>;
    _readSchemaExtension(extensionName: string): Promise<any>;
    _updateSchemaExtension(extensionName: string, meTaskSettings: string): Promise<any>;
    _deleteSchemaExtension(extensionName: string): Promise<any>;
}

export default async function msGraphProvider(msGraphClientFactory: MSGraphClientFactory): Promise<IGraphServices> {
    const msGraphClient = await msGraphClientFactory.getClient();

    const _createSchemaExtension = async (extensionName: string, favoritePlane: string) => {

        let _extensionResult: any;
        let extentionData: Object = {};
        try {
            extentionData = {
                "@odata.type": "#microsoft.graph.openTypeExtension",
                extensionName: extensionName,
                plane: favoritePlane
            };
            _extensionResult = await msGraphClient.api(`/me/extensions`).post(extentionData);
        } catch (ex) {
            console.log(ex);
        }
        return _extensionResult;
    };

    const _readSchemaExtension = async (extensionName: string) => {

        let _extensionResult: any;
        try {
            _extensionResult = await msGraphClient.api(`/me/extensions/${extensionName}`).get();
        } catch (ex) {
            console.log(ex);
        }
        return _extensionResult;
    };

    const _updateSchemaExtension = async (extensionName: string, favoritePlane: string) => {

        let _extensionResult: any;
        let extentionData: Object = {};
        try {
            extentionData = {
                "@odata.type": "#microsoft.graph.openTypeExtension",
                extensionName: extensionName,
                plane: favoritePlane
            };
            _extensionResult = await msGraphClient.api(`/me/extensions/${extensionName}`).patch(extentionData);

        } catch (ex) {
            console.log(ex);
        }
        return _extensionResult;
    };

    const _deleteSchemaExtension = async (extensionName: string) => {

        let _extensionResult: any;
        try {
            _extensionResult = await msGraphClient.api(`/me/extensions/${extensionName}`).delete();
        } catch (ex) {
            console.log(ex);
        }
        return _extensionResult;
    };

    return {
        _createSchemaExtension,
        _readSchemaExtension,
        _updateSchemaExtension,
        _deleteSchemaExtension
    };

}