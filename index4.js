const xlsx = require('xlsx');
let wb = xlsx.readFile('./tastemade_ids-platform.xlsx');
const fs = require('fs');

fs.readFile('./data.json', 'utf8', (err, fileContents) => {
    if (err) {
        console.error(err)
        return err;
    }
    try {
        const Contents = JSON.parse(fileContents);
        var main_element = Contents['rss']['series'];
        for (var element in main_element) {
            var main_element_item = main_element[element]['item']
        }
        let itemFromepisodes = [];
        main_element_item.forEach(function (arrayItem) {
            itemFromepisodes.push(arrayItem.episodes);

        });
        var itemFromepisodes_filtered = itemFromepisodes.filter(n => n)

        function getFields(input) {
            var output = [];

            for (var i = 0; i < input.length; i++)
                output.push(input[i][0]);
            return output;
        }
        var resultfin = getFields(itemFromepisodes_filtered);

        let res = [];
        resultfin.forEach((mass) => {
            res.push(mass.item)
        })

        var merged = [].concat.apply([], res);
        let resto = [];
        merged.forEach((element) => {
            if (!element?.guid[0]?._) return;

            let temporary = element['guid'][0]._
            let finalArray = [];
            (element['tastemade:meta'][0]['tastemade:meta-list'] || []).map((res) => {
                if (res.$.name == 'svodPlatformWhitelist') {
                    (res['tastemade:meta-value'] || []).map((obj, i) => {
                        let valuesOfPlatform = (obj.$.value) ? obj.$.value : 'null';
                        finalArray.push(valuesOfPlatform);
                    })
                }
            });
            resto.push({ "id": temporary, "platform": `${finalArray}` });
        }
        )
        let workSheets = {};
        for (const sheetName of wb.SheetNames) {
            workSheets[sheetName] = xlsx.utils.sheet_to_json(wb.Sheets[sheetName])
        }
        let Guids = workSheets.Sheet1;


        const finalData = (Guids || []).map((g) => {
            const r = resto.find((r) => r.id === g.episode_guid);
            return {
                id: (r && r.id) || g.episode_guid,
                platform: (r && r.platform) || null
            }
        });

        const newBook = xlsx.utils.book_new();
        const newSheet = xlsx.utils.json_to_sheet(finalData);
        xlsx.utils.book_append_sheet(newBook, newSheet, "output");
        xlsx.writeFile(newBook, "outputnew.xlsx")
        console.log("done!!!");
    } catch (err) {
        console.error(err)
    }
})
;