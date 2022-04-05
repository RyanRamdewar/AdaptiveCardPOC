import { IInputs } from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;

export interface IGridProps {
    context: ComponentFramework.Context<IInputs>,
    columns: DataSetInterfaces.Column[] | undefined,
    rows: ComponentFramework.PropertyTypes.DataSet;
    entityId: string,
    entityType: string,
}

export interface IContact {
    fName: string,
    lName: string,
    email: string

}