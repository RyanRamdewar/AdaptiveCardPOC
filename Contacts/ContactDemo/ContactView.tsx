import { DetailsList, DetailsListLayoutMode, IColumn, PrimaryButton, SelectionMode } from "@fluentui/react";
import { AdaptiveCard } from "adaptivecards-react";
import { useState } from "react";
import React = require("react");
import { IGridProps, IContact } from "./Interfaces";

const ContactView = (gridProps: IGridProps) => {
    const [items, setItems] = useState<IContact[]>([]);

    React.useEffect(() => {
        refreshItems(gridProps);
    }, [gridProps.rows.records])

    function refreshItems(gridProps: IGridProps) {
        let i: IContact[] = [];
        for (const recordId in gridProps.rows.records) {
            const rec = gridProps.rows.records[recordId];
            const fName = rec.getFormattedValue("firstname");
            const lName = rec.getFormattedValue("lastname");
            const email = rec.getFormattedValue("emailaddress1");
            i.push({
                fName,
                lName,
                email
            });
        }
        setItems(i);
    }

    const columns: IColumn[] = [
        {
            key: 'column1',
            name: 'First Name',
            fieldName: 'fName',
            maxWidth: 150,
            minWidth: 150,
            isRowHeader: true,
            isResizable: true,
            isSorted: true,
            isSortedDescending: false,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            data: 'string',
            isPadded: true,

        },
        {
            key: 'column2',
            name: 'Last Name',
            fieldName: 'lName',
            minWidth: 150,
            maxWidth: 150,
            isRowHeader: true,
            isResizable: true,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            data: 'string',
            isPadded: true,
        },
        {
            key: 'column4',
            name: 'E-Mail',
            fieldName: 'email',
            minWidth: 150,
            maxWidth: 150,
            isRowHeader: true,
            isResizable: true,
            sortAscendingAriaLabel: 'Sorted Ascending',
            sortDescendingAriaLabel: 'Sorted Descending',
            data: 'string',
            //isPadded: true,
        },


    ];
    var card = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.3",
        "body": [
            {
                "type": "Input.Text",
                "placeholder": "First Name",
                "id": "fName"
            },
            {
                "type": "Input.Text",
                "placeholder": "Last Name",
                "id": "lName"
            },
            {
                "type": "Input.Text",
                "placeholder": "E-Mail",
                "id": "email"
            },
            {
                "type": "ActionSet",
                "actions": [
                    {
                        "type": "Action.Submit",
                        "title": "Submit"
                    }
                ],
                "id": "submit"
            }
        ]
    };
    function onSubmit(a: any) {
        var record: any = {};
        record["parentcustomerid_account@odata.bind"] = "/accounts(" + gridProps.entityId + ")"; // Customer
        record.firstname = a.data.fName; // Text
        record.lastname = a.data.lName; // Text
        record.emailaddress1 = a.data.email; // Text
        gridProps.context.webAPI.createRecord("contact", record).then(() => {
            gridProps.rows.refresh();
            items.push({ fName: a.data.fName, lName: a.data.lName, email: a.data.email })
            setItems(items);
        })
    }
    return <React.Fragment>
        <div style={{ maxHeight: "315px", overflowY: 'auto' }}>
            <DetailsList
                items={items}
                compact={true}
                columns={columns}
                selectionMode={SelectionMode.single}
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                // onItemInvoked={_onItemInvoked}
                enterModalSelectionOnTouch={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                checkButtonAriaLabel="Row checkbox"
            />
        </div>
    </React.Fragment>
}

export default ContactView;

