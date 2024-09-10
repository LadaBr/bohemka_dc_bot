const { Client, Intents } = require('discord.js');
const { GoogleSpreadsheet } = require('google-spreadsheet');

const creds = require('./still-lamp-327420-417d08612518.json');
const moment = require('moment-timezone');
moment.locale('cs');
moment().tz("Europe/Prague").format();
const client = new Client({
    intents: [Intents.FLAGS.GUILDS, Intents.FLAGS.GUILD_MESSAGES, Intents.FLAGS.GUILD_MESSAGE_REACTIONS],
    partials: ['MESSAGE', 'CHANNEL', 'REACTION'],
});
const channelId = "843912837033361409";
const botId = "579155972115660803";
let sheet;
const NUM_GROUPS = 7;
const RANGE = 'M3:W32';
const NUM_GROUP_MEMBERS = 5;
const specs = {
    "Restoration1": " Restoration",
    "Protection1": "Protection",
    "Protection": " Protection",
    "Guardian": "Feral",
    "Holy1": "Holy ",
};
const classColors = {
    "Shaman": "#0057ac",
    "Druid": "#FF7D0A",
    "Hunter": "#ABD473",
    "Mage": "#69CCF0",
    "Paladin": "#F58CBA",
    "Priest": "#FFFFFF",
    "Rogue": "#FFF569",
    "Warrior": "#C79C6E",
    "Warlock": "#9482C9",
};
const specToClass = {
    "Protection": "Warrior",
    "Protection1": "Paladin",
    "Guardian": "Druid",
};
const DATE_FORMAT = 'DD-MM-YYYY HH:mm:ss';

function hexToRgb(hex) {
    var shorthandRegex = /^#?([a-f\d])([a-f\d])([a-f\d])$/i;
    hex = hex.replace(shorthandRegex, function(m, r, g, b) {
        return r + r + g + g + b + b;
    });

    var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        red: parseInt(result[1], 16) / 255,
        green: parseInt(result[2], 16) / 255,
        blue: parseInt(result[3], 16) / 255
    } : null;
}

client.on('ready', async () => {
    console.log(`Logged in as ${client.user.tag}!`);
    // Initialize the sheet - doc ID is the long id in the sheets URL
    const doc = new GoogleSpreadsheet('1aRdpJOPfhayPlEGcDYl9PshM2BRiIs3zjfGTdesNBtg');

    // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
    await doc.useServiceAccountAuth(creds);

    await doc.loadInfo(); // loads document properties and worksheets
    sheet = doc.sheetsByIndex[0];
    const raids = await loadRaids();
    clearOldRaids(raids);
    await sheet.saveUpdatedCells();
});


client.on('messageUpdate', async (oldMessage, newMessage) => {
    const data = {
        players: [],
        count: {}
    };
    if (newMessage.author.id === botId) {
        if (channelId !== newMessage.channelId) return;
        data.id = newMessage.id;
        for (const embed of newMessage.embeds) {
            const desc = embed.description.match(/<:(.*?):/gm);
            let name = "";
            for (const d of desc) {
                if (d.includes("empty")) {
                    name += " ";
                } else if (d.includes("plus")) {
                    name += "+";
                } else if (d.includes("minus")) {
                    name += "-";
                } else {
                    name += d.substring(2, 3);
                }
            }
            data.name = name;
            for (const field of embed.fields) {
                const rows = field.value.split("\n");
                let rowType;
                let className;
                for (const [index, value] of rows.entries()) {
                    rowType = value.match(/<:(.*?):/);

                    if (rowType && rowType.length === 2) {
                        rowType = rowType[1];
                    }
                    if (rowType === "CMcalendar") {
                        let date = value.match(/dates=(.+?)&/);
                        if (date && date.length > 1) {
                            let dateTime = decodeURIComponent(date[1]).split("/")[0];
                            dateTime = dateTime.split("T");
                            date = dateTime[0];
                            let time = dateTime[1];
                            const year = date.substring(0, 4);
                            const month = date.substring(4, 6);
                            const day = date.substring(6, 8);
                            time = time.match(/\d{1,2}/g).join(":");
                            date = moment.utc(new Date(`${year}-${month}-${day}T${time}Z`)).tz("Europe/Prague");
                            data.date = date;
                        }
                    } else if (rowType === "Tanks" || rowType === "Ranged" || rowType === "Healers") {
                        let count;
                        if (rowType === "Tanks") {
                            count = value.match(/\*\* (\d+?)\*\*/);
                            if (count.length) {
                                data.count[rowType] = parseInt(count[1]);
                            }
                            count = value.match(/\*\*(\d+?)\*\*/);
                            if (count.length) {
                                data.count["Melee"] = parseInt(count[1]);
                            }
                        } else {
                            count = value.match(/\*\*(\d+?)\*\*/);
                            if (count.length) {
                                data.count[rowType] = parseInt(count[1]);
                            }
                        }

                    } else if (rowType) {
                        if (index === 0) {
                            className = rowType;
                        }
                        let player = value.match(/<:(.*?):.*`(\d+)`.*\*\*(.+)\*\*/);
                        if (player && player.length > 3) {
                            const spec = player[1];
                            if (spec === "Bench" || spec === "Absence" || spec === "Late") continue;
                            const id = parseInt(player[2]);
                            const playerName = player[3];
                            data.players.push({
                                className: className === "Tank" ? specToClass[spec] : className,
                                spec,
                                id,
                                playerName
                            });
                        }
                    }


                }
            }
        }
        data.players.sort((a, b) => a.id - b.id);

        const raids = await loadRaids();
        clearOldRaids(raids);
        let existingRaid = raids.find(r => r.cells && r.cells.id.value === data.id);
        if (!existingRaid) {
            console.log("Creating new raid", data.id);
            existingRaid = raids.find(r => !r.cells.id.value);
            if (!existingRaid) {
                console.log("full slots");
                return;
            }
        } else {
            const index = raids.indexOf(existingRaid);
            for (let i = 0; i < index; i++) {
                const raid = raids[i];
                if (raid && !raid.cells.id.value) {
                    clearRaid(existingRaid);
                    existingRaid = raid;
                }
            }
        }
        const cells = existingRaid.cells;
        cells.id.value = data.id;
        cells.date.value = data.date.format(DATE_FORMAT);
        cells.isoDate.value = data.date.format();
        cells.name.value = data.name;
        const raidComp = getGroups(sheet, "O3");
        clearRaidRoster(existingRaid);
        await sheet.saveUpdatedCells();
        const first = sheet.getCellByA1(existingRaid.groups);
        for (const [index, [name, count]] of Object.entries(data.count).entries()) {
            existingRaid.cells.counts[index].name.value = name;
            existingRaid.cells.counts[index].value.value = count;
        }
        for (const slot of raidComp) {
            for (let i = 0; i < data.players.length; i++) {
                const player = data.players[i];
                if (!player) break;
                const spec = specs[player.spec] || player.spec;
                if (!spec) continue;
                if (slot.cell.value === spec || slot.j > 4) {
                    const cell = sheet.getCell(first.rowIndex + slot.i, first.columnIndex + slot.j);
                    const isValidClass = classColors[player.className];
                    if (!cell.value && isValidClass) {
                        cell.value = player.playerName;
                        cell.backgroundColor = hexToRgb(isValidClass);
                        data.players.splice(i, 1);
                        break;
                    }
                }
            }
        }

        await sheet.saveUpdatedCells();
    }
});

const loadRaids = async () => {
    const raids = [
        {
            date: 'R10:S10',
            name: 'O10:P10',
            groups: 'O12',
            id: 'T10',
            isoDate: 'U10',
            counts: 'M13',
        },
        {
            date: 'R18:S18',
            name: 'O18:P18',
            groups: 'O20',
            id: 'T18',
            isoDate: 'U18',
            counts: 'M21',
        },
        {
            date: 'R26:S26',
            name: 'O26:P26',
            id: 'T26',
            groups: 'O28',
            isoDate: 'U26',
            counts: 'M29',
        },
    ];
    await sheet.loadCells(RANGE);
    loadRaidDataFromSheet(raids);
    return raids;
}
const clearOldRaids = (raids) => {
    for (const raid of raids) {
        if (raid.cells.date.value && (moment() - moment(raid.cells.isoDate.value)) > 0) {
            console.log("Deleting old raid", raid.cells.id.value);
            clearRaid(raid);
        }
    }
}
const clearRaid = (raid) => {
    clearRaidRoster(raid);
    raid.cells.id.value = "";
    raid.cells.name.value = "";
    raid.cells.date.value = "";
    raid.cells.isoDate.value = "";
    for (const cell of raid.cells.counts) {
        cell.name.value = "";
        cell.value.value = "";
    }
    delete raid.cells;
}
const clearRaidRoster = (raid) => {
    for (const cell of raid.cells.groups) {
        cell.cell.backgroundColor = hexToRgb(cell.j > 4 ? "#2a2a2a": "#000000");
        cell.cell.value = "";
    }
}
const loadRaidDataFromSheet = (raids) => {
    for (const raid of raids) {
        const date = sheet.getCellByA1(raid.date);
        const name = sheet.getCellByA1(raid.name);
        const groups = getGroups(sheet, raid.groups);
        const id = sheet.getCellByA1(raid.id);
        const isoDate = sheet.getCellByA1(raid.isoDate);
        raid.cells = {
            date,
            name,
            groups,
            id,
            isoDate,
            counts: getCounts(sheet, raid.counts),
        }
    }
}

const getGroups = (sheet, start) => {
    const first = sheet.getCellByA1(start);
    const groups = [];
    for (let j = 0; j < NUM_GROUPS;j++) {
        for (let i = 0; i < NUM_GROUP_MEMBERS;i++) {
            const cell = sheet.getCell(first.rowIndex + i, first.columnIndex + j);
            groups.push({cell, i, j});
        }
    }
    return groups;
}

const getCounts = (sheet, start) => {
    const first = sheet.getCellByA1(start);
    const groups = [];
    for (let j = 0; j < 4;j++) {
        const cell = sheet.getCell(first.rowIndex + j, first.columnIndex);
        const cell2 = sheet.getCell(first.rowIndex + j, first.columnIndex + 1);
        groups.push({name: cell, value: cell2});
    }
    return groups;
}


(async () => {
    client.login('');
})().catch(e => console.error(e));



