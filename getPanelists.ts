
import * as jwt from 'jsonwebtoken';
import * as rp from 'request-promise';
import * as xlsx from 'xlsx';


const payload = {
	iss: 'ключ',
	exp: ((new Date()).getTime() + 5000)
};

const token = jwt.sign(payload, 'секрет');

const options = {
	uri: "https://api.zoom.us/v2/webinars/89888489318/panelists",
	auth: {
		'bearer': token
	},
	headers: {
		'User-Agent': 'Zoom-api-Jwt-Request',
		'content-type': 'application/json'
	},
	json: true
};

type PanelistResponseItem = {
	id: string;
	name: string;
	email: string;
	join_url: string;
};

type PanelistResponse = {
	panelists: PanelistResponseItem[];
};

rp(options)
	.then((response: PanelistResponse) => {
		const filename = "panelists.xlsx";
		const data = response.panelists.map(p => [p.name, p.email, p.join_url]);
		const wsName = "Выступающие";

		const wb = xlsx.utils.book_new();
		const ws = xlsx.utils.aoa_to_sheet(data);
		xlsx.utils.book_append_sheet(wb, ws, wsName);
		xlsx.writeFile(wb, filename);
	});