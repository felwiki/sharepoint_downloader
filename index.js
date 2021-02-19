const puppeteer = require('puppeteer');
const cli_progress = require('cli-progress');
const path = require('path');
const fs = require('fs');
const commander = require('commander');

const program = new commander.Command();
program.version('0.1.0');
program.requiredOption('-i, --input <url>', 'Url address of sharepoint file');
program.requiredOption('-u, --username <username>', 'ČVUT username (without domain)');
program.requiredOption('-p, --password <password>', 'Password to the account');
program.requiredOption('-o, --output <filename>', 'Downloaded file path');
program.option('--tmp <path>', 'Path tmp directory', '/tmp/');

program.parse(process.argv);
const options = program.opts();

async function cvut_login(browser, page, username, password) {
	console.log("Logging to sharepoint...");

	await page.goto('https://campuscvut.sharepoint.com/', { waitUntil: 'load' });

	await page.waitForSelector('input[type="email"]', { timeout: 3000 });
	await page.keyboard.type(username + "@cvut.cz");
	await page.click('input[type="submit"]');
	console.log("Username filled");

	await browser.waitForTarget((target) => target.url().startsWith('https://logon.ms.cvut.cz'), { timeout: 15000 });
	await page.waitForSelector('input[type="password"]', { timeout: 3000 });
	await page.keyboard.type(password);
	await page.click('#submitButton');
	console.log("Password filled");

	await browser.waitForTarget((target) => target.url().startsWith('https://login.microsoftonline.com/'), { timeout: 15000 });
	await page.waitForSelector('input[type="submit"]', { timeout: 3000 });
	await page.click('input[type="submit"]');
	console.log("Autologin skipped");

	await browser.waitForTarget((target) => target.url().startsWith('https://campuscvut.sharepoint.com/'), { timeout: 15000 });
	console.log("Login successful");
}

var progress_bar = null;

async function sharepoint_download(browser, page, url, tmp_path) {
	var cdp = await page.target().createCDPSession();
	await cdp.send('Browser.setDownloadBehavior', {
		behavior: 'allowAndName',
		downloadPath: tmp_path
	});

	await cdp.send('Page.enable');
	let download_promise = new Promise((resolve, reject) => {
		cdp.on('Page.downloadWillBegin', (data) => {
			console.log("Download begin " + path.join(tmp_path, data.guid)); progress_bar = new cli_progress.SingleBar({ format: "{bar} | ETA: {eta}s | {percentage}%" }, cli_progress.Presets.shades_classic);
			progress_bar.start(1, 0);
		});

		cdp.on('Page.downloadProgress', (data) => {
			progress_bar.update(data.receivedBytes / data.totalBytes);

			if (data.state === 'completed') {
				progress_bar.stop();
				resolve(path.join(tmp_path, data.guid));
			} else if (data.state === 'canceled') {
				reject(data);
			}
		});
	});

	console.log("Navigate to sharepoint website");
	await page.goto(url, { waitUntil: 'load' });

	await page.waitForSelector('button[data-automationid="download"]', { timeout: 3000 });
	await page.click('button[data-automationid="download"]');

	return await download_promise;
}

async function download() {
	const browser = await puppeteer.launch({
		headless: true,
		userDataDir: undefined,
		devtools: true,
		args: [
			'--disable-dev-shm-usage',
			'--fast-start',
			'--no-sandbox'
		]
	});

	const page = (await browser.pages())[0];
	await cvut_login(browser, page, options.username, options.password);
	let download_path = await sharepoint_download(browser, page, options.input, options.tmp);
	await browser.close();

	fs.mkdirSync(path.dirname(options.output), { recursive: true });
	fs.copyFileSync(download_path, options.output);
	fs.rmSync(download_path);

	console.log("Downloaded " + options.output);
}

download().catch((err) => {
	console.error(err);
	process.exit(1);
});