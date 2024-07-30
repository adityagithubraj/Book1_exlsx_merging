const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });


app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.render('index');
});

app.post('/merge', upload.fields([{ name: 'book1' }, { name: 'book2' }]), (req, res) => {
    try {
        const book1Path = req.files['book1'][0].path;
        const book2Path = req.files['book2'][0].path;

        const book1 = xlsx.readFile(book1Path);
        const book2 = xlsx.readFile(book2Path);

        const sheet1 = book1.Sheets[book1.SheetNames[0]];
        const sheet2 = book2.Sheets[book2.SheetNames[0]];

        const data1 = xlsx.utils.sheet_to_json(sheet1);
        const data2 = xlsx.utils.sheet_to_json(sheet2);

        // Create a mapping of pincode to codes from book1
        const codeMap = data1.reduce((map, row) => {
            const pincode = row['PIN'];
            if (!map[pincode]) {
                map[pincode] = [];
            }
            map[pincode].push(row.POD);
            return map;
        }, {});

        // Apply the mapping to book2 data
        const mergedData = data2.map(row => {
            const pincode = row['PIN'];
            if (codeMap[pincode] && codeMap[pincode].length > 0) {
                return {
                    ...row,
                    code: codeMap[pincode].shift()
                };
            }
            return row;
        });

        const newSheet = xlsx.utils.json_to_sheet(mergedData);
        book2.Sheets[book2.SheetNames[0]] = newSheet;

        const outputDir = path.join(__dirname, 'final_data');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
        }

        const outputFilePath = path.join(outputDir, 'mergedBook.xlsx');
        xlsx.writeFile(book2, outputFilePath);

        res.download(outputFilePath, 'mergedBook.xlsx', (err) => {
            if (err) {
                console.error(err);
                res.status(500).send('Error occurred while sending the file');
            }

            // Cleanup temporary files
            fs.unlinkSync(book1Path);
            fs.unlinkSync(book2Path);
            // fs.unlinkSync(outputFilePath); // Commented out to keep the merged file in final_data
        });
    } catch (error) {
        console.error(error);
        res.status(500).send('An error occurred while merging the files');
    }
});

const PORT = process.env.PORT || 8000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
