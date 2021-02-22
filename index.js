const puppeteer = require('puppeteer');
const cli_progress = require('cli-progress');
const path = require('path');
const fs = require('fs');
const commander = require('commander');

const LOGIN_ERROR_CODE = 7; //ERROR_CODE.NO_SESSION_INFO from destreamer
const SELECTOR_TIMEOUT = 15000;
const TARGET_TIMEOUT = 30000;

const program = new commander.Command();
program.version('0.1.0');
program.requiredOption('-i, --input <url>', 'Url address of sharepoint file');
program.requiredOption('-o, --output <filename>', 'Downloaded file path');
program.option('-u, --username <username>', 'ČVUT username (without domain)');
program.option('-p, --password <password>', 'Password to the account');
program.option('--tmp <path>', 'Path to tmp directory', '/tmp/');
program.option('--chromeData <path>', 'Path to cache directory', '.chrome_data');

program.parse(process.argv);
const options = program.opts();

async function cvut_login(browser, page, username, password) {
	console.log("Logging to sharepoint...");

	await page.goto('https://campuscvut.sharepoint.com/', { waitUntil: 'load' });

	if(!page.url().startsWith('https://login.microsoftonline.com')) {
		console.log("User already logined");
		return;
	}

	if(!username || !password) {
		console.error("Invalid login credentials");
		process.exit(LOGIN_ERROR_CODE);
	}

	try {
		await page.waitForSelector('input[type="email"]', { timeout: SELECTOR_TIMEOUT });
		await page.keyboard.type(username + "@cvut.cz");
		await page.click('input[type="submit"]');
		console.log("Username filled");

		await browser.waitForTarget((target) => target.url().startsWith('https://logon.ms.cvut.cz'), { timeout: TARGET_TIMEOUT });
		await page.waitForSelector('input[type="password"]', { timeout: SELECTOR_TIMEOUT });
		await page.keyboard.type(password);
		await page.click('#submitButton');
		console.log("Password filled");

		await browser.waitForTarget((target) => target.url().startsWith('https://login.microsoftonline.com/'), { timeout: TARGET_TIMEOUT });
		await page.waitForSelector('input[type="submit"]', { timeout: SELECTOR_TIMEOUT });
		await page.click('input[type="submit"]');
		console.log("Autologin skipped");

		await browser.waitForTarget((target) => target.url().startsWith('https://campuscvut.sharepoint.com/'), { timeout: TARGET_TIMEOUT });
	} catch (err) {
		console.error("Invalid login");
		process.exit(LOGIN_ERROR_CODE);
	}
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

	await page.waitForSelector('button[data-automationid="download"]', { timeout: SELECTOR_TIMEOUT });
	await page.click('button[data-automationid="download"]');

	return await download_promise;
}

async function download() {
	const browser = await puppeteer.launch({
		headless: true,
		userDataDir: options.chromeData,
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
	fs.unlinkSync(download_path);

	console.log("Downloaded " + options.output);
}

download().catch((err) => {
	console.error(err);
	process.exit(1);
});