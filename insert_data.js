const xlsx = require('xlsx');
const { MongoClient } = require('mongodb');

async function main() {
    const uri = 'mongodb+srv://gadalRentAdmin:wtLYPxGPKcd9hVfs@gadalrent.yugiznm.mongodb.net/?retryWrites=true&w=majority';
    const client = new MongoClient(uri, { useNewUrlParser: true, useUnifiedTopology: true });

    try {
        await client.connect();
        const db = client.db('Gadal_Market2');

        // Read Excel file
        const workbook = xlsx.readFile('data.xlsx'); // Ensure the data.xlsx is in the same directory as this script

        // Get single sheet
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // Convert sheet to JSON
        const data = xlsx.utils.sheet_to_json(sheet);

        // Collections
        const locationCollection = db.collection('locations');
        const subcityCollection = db.collection('subcities');
        const woredaCollection = db.collection('weredas');

        // Maps to store generated IDs
        const locationMap = {};
        const subcityMap = {};
        var i = 0;
        // Insert data
        for (const row of data) {
           i++;
           console.log(i)
            // Insert location if not already inserted
            if (!locationMap[row.LocationName]) {
                const locationResult = await locationCollection.insertOne({ descripton: row.LocationName,recordStatus:1 });
                locationMap[row.LocationName] = locationResult.insertedId;
            }

            // Insert subcity if not already inserted
            if (!subcityMap[row.SubcityName]) {
                const subcityResult = await subcityCollection.insertOne({
                    discripton: row.SubcityName,
                    location: locationMap[row.LocationName],
                    recordStatus:1

                });
                subcityMap[row.SubcityName] = subcityResult.insertedId;
            }

            // Insert woreda
            await woredaCollection.insertOne({
                discripton: row.WoredaName,
                subCity: subcityMap[row.SubcityName],
                recordStatus:1
            });
        }

        console.log('Data inserted successfully');
    } finally {
        await client.close();
    }
}

main().catch(console.error);
