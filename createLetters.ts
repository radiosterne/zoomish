
import * as xlsx from 'xlsx';

const inputWorkbook = xlsx.readFile('panelists-template.xlsx');
const inputWorksheet = inputWorkbook.Sheets[inputWorkbook.SheetNames[0]];
const inputArray = xlsx.utils.sheet_to_json(inputWorksheet, { header: 1 }) as string[][];

const template = `Добрый день!

[0], приглашаю Вас принять участие в zoom-вебинаре Челябинского Тракторного завода.

Вы выступаете в [1], ваше выступление начинается после выступления [2].

Ваша ссылка на вход — [3].

С уважением, PR-служба Челябинского тракторного.`

const substitute = (data: string[]) =>
	data.reduce((currentTemplate, currentValue, currentIndex) => currentTemplate.replace(`[${currentIndex}]`, currentValue), template)

const result = inputArray
	.map(value => substitute(value))
	.map(substituted => [substituted]);

const filename = "letters.xlsx";
const wsName = "Письма";

const wb = xlsx.utils.book_new();
const ws = xlsx.utils.aoa_to_sheet(result);
xlsx.utils.book_append_sheet(wb, ws, wsName);
xlsx.writeFile(wb, filename);