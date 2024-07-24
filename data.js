import { appendValues, updateValues, getValues } from '@velo/google-sheets-integration-backend';
import { getSecret } from 'wix-secrets-backend';

// MembersInformation Hooks
export async function MembersInformation_afterInsert(item, context) {
    try {
        const sheetId = await getSecret("sheetId");
        await syncMemberToGoogleSheet(item, 'append', sheetId);
    } catch (error) {
        console.error("Error in MembersInformation_afterInsert:", error);
    }
    return item;
}

export async function MembersInformation_afterUpdate(item, context) {
    try {
        const sheetId = await getSecret("sheetId");
        await syncMemberToGoogleSheet(item, 'update', sheetId);
    } catch (error) {
        console.error("Error in MembersInformation_afterUpdate:", error);
    }
    return item;
}

// Companies Hooks
export async function Companies_afterInsert(item, context) {
    try {
        const sheetId = await getSecret("sheetId");
        await syncCompanyToGoogleSheet(item, 'append', sheetId);
    } catch (error) {
        console.error("Error in Companies_afterInsert:", error);
    }
    return item;
}

export async function Companies_afterUpdate(item, context) {
    try {
        const sheetId = await getSecret("sheetId");
        await syncCompanyToGoogleSheet(item, 'update', sheetId);
    } catch (error) {
        console.error("Error in Companies_afterUpdate:", error);
    }
    return item;
}

// Synchronization Functions
async function syncMemberToGoogleSheet(member, operation, sheetId) {
    const sheetName = 'Users';
    const rows = await getRows(sheetId, sheetName);
    const userColumns = createUserColumnMapping(); // Use existing column mapping
    const sheetColumns = createColumnMapping(rows[0]); // Create column mapping based on header row
    const values = mapValuesToColumns(userColumns, sheetColumns, member);

    console.log('Values to be sent:', values);

    if (operation === 'append') {
        await appendValues(sheetId, `${sheetName}!A:ZZ`, [values]);
    } else if (operation === 'update') {
        const rowIndex = await getRowIndexByEmail(rows, sheetColumns, member.emailAddress);
        if (rowIndex) {
            console.log('Updating row:', rowIndex, 'with values:', values);
            await updateSpecificCells(sheetId, sheetName, rowIndex, values, sheetColumns);
        } else {
            // If the entry does not exist in Google Sheet, insert it
            await appendValues(sheetId, `${sheetName}!A:ZZ`, [values]);
        }
    }
}

async function syncCompanyToGoogleSheet(company, operation, sheetId) {
    const sheetName = 'Companies';
    const rows = await getRows(sheetId, sheetName);
    const companyColumns = createCompanyColumnMapping(); // Use existing column mapping
    const sheetColumns = createColumnMapping(rows[0]); // Create column mapping based on header row
    const values = mapValuesToColumns(companyColumns, sheetColumns, company);

    console.log('Values to be sent:', values);

    if (operation === 'append') {
        await appendValues(sheetId, `${sheetName}!A:ZZ`, [values]);
    } else if (operation === 'update') {
        const rowIndex = await getRowIndexByCompanyId(rows, sheetColumns, company.company_id);
        if (rowIndex) {
            console.log('Updating row:', rowIndex, 'with values:', values);
            await updateSpecificCells(sheetId, sheetName, rowIndex, values, sheetColumns);
        } else {
            // If the entry does not exist in Google Sheet, insert it
            await appendValues(sheetId, `${sheetName}!A:ZZ`, [values]);
        }
    }
}

// Helper Functions
async function getRows(sheetId, sheetName) {
    try {
        const response = await getValues(sheetId, `${sheetName}!A:ZZ`);
        const rows = response.data.values;
        if (!rows || rows.length === 0) {
            console.log("Rows are empty or null.");
            return [];
        }
        console.log("Fetched rows:", rows);
        return rows;
    } catch (error) {
        console.error("Error fetching rows:", error);
        return [];
    }
}

async function getRowIndexByEmail(rows, columns, email) {
    const emailColumnIndex = columns['email'];
    if (emailColumnIndex === undefined) return null;
    const rowIndex = rows.findIndex(row => row[emailColumnIndex] === email);
    if (rowIndex === -1) {
        console.log(`Email ${email} not found.`);
        return null;
    }
    return rowIndex + 1; // Add 1 to convert to 1-based index
}

async function getRowIndexByCompanyId(rows, columns, companyId) {
    const companyIdColumnIndex = columns['company_id'];
    if (companyIdColumnIndex === undefined) return null;
    const rowIndex = rows.findIndex(row => row[companyIdColumnIndex] === companyId);
    if (rowIndex === -1) {
        console.log(`Company ID ${companyId} not found.`);
        return null;
    }
    return rowIndex + 1; // Add 1 to convert to 1-based index
}

function createUserColumnMapping() {
    return {
        emailAddress: 'email',
        user_id: 'user_id',
        name: 'first_name',
        lastName: 'last_name',
        contact: 'primary_phone',
        personalDescription: 'personalDescription',
        companyDescription: 'companyDescription',
        location: 'location',
        preferred_first_name: 'preferred_first_name',
        middleName: 'middle_name',
        secondaryEmail: 'secondary_email',
        sex: 'sex',
        referred_by: 'referred_by',
        email_verified: 'verification_status'
    };
}

function createCompanyColumnMapping() {
    return {
        company_name: 'company_name',
        company_address: 'company_address',
        mailing_address_optional: 'mailing_address_optional',
        company_category: 'company_category',
        company_email: 'company_email',
        company_phone: 'company_phone',
        fax_company_phone_optional: 'fax_company_phone_optional',
        website_url: 'website_url',
        social_media: 'social_media',
        company_about: 'company_about',
        active_corporate_membership: 'active_corporate_membership',
        active_corporate_membership_type: 'active_corporate_membership_type',
        rep_1: 'rep_1',
        rep_2: 'rep_2',
        rep_3: 'rep_3'
    };
}

function createColumnMapping(headerRow) {
    const columnMapping = {};
    headerRow.forEach((columnName, index) => {
        columnMapping[columnName] = index;
    });
    return columnMapping;
}

function mapValuesToColumns(columnMapping, sheetColumns, data) {
    const orderedValues = new Array(Object.keys(sheetColumns).length).fill('');
    for (const [key, value] of Object.entries(columnMapping)) {
        if (sheetColumns[value] !== undefined) {
            orderedValues[sheetColumns[value]] = data[key] || '';
        }
    }
    return orderedValues;
}

async function updateSpecificCells(sheetId, sheetName, rowIndex, values, sheetColumns) {
    const updates = [];

    for (const [columnName, columnIndex] of Object.entries(sheetColumns)) {
        if (values[columnIndex] !== '') {
            updates.push({
                range: `${sheetName}!${String.fromCharCode(65 + columnIndex)}${rowIndex}`,
                values: [[values[columnIndex]]]
            });
        }
    }

    const requests = updates.map(update => updateValues(sheetId, update.values, update.range, "ROWS"));

    await Promise.all(requests);
}
