#!/usr/bin/env node
// formflow-cli.js - Terminal-based Form Filler untuk Linux
// Install: npm install puppeteer xlsx readline-sync chalk ora

const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const readline = require('readline-sync');
const chalk = require('chalk');
const ora = require('ora');
const fs = require('fs');
const path = require('path');

// ========== CONFIGURATION ==========
const CONFIG = {
    headless: false, // Set true untuk mode headless
    timeout: 30000,
    otpDelay: 2000,
    errorResetDelay: 3000,
    viewport: { width: 430, height: 932 } // Mobile viewport
};

// ========== STATE MANAGEMENT ==========
class FormFlow {
    constructor() {
        this.browser = null;
        this.page = null;
        this.data = {
            headers: [],
            rows: [],
            mappings: [],
            history: this.loadHistory(),
            targetUrl: ''
        };
    }

    // Load history dari file JSON
    loadHistory() {
        const historyFile = path.join(__dirname, 'formflow_history.json');
        try {
            if (fs.existsSync(historyFile)) {
                return JSON.parse(fs.readFileSync(historyFile, 'utf8'));
            }
        } catch (e) {
            console.error(chalk.red('Error loading history:', e.message));
        }
        return [];
    }

    // Save history ke file JSON
    saveHistory() {
        const historyFile = path.join(__dirname, 'formflow_history.json');
        try {
            fs.writeFileSync(historyFile, JSON.stringify(this.data.history, null, 2));
        } catch (e) {
            console.error(chalk.red('Error saving history:', e.message));
        }
    }

    // ========== FILE HANDLING ==========
    async loadExcelFile(filePath) {
        const spinner = ora('Membaca file Excel...').start();
        
        try {
            if (!fs.existsSync(filePath)) {
                spinner.fail(chalk.red('File tidak ditemukan!'));
                return false;
            }

            const workbook = XLSX.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

            if (data.length === 0) {
                spinner.fail(chalk.red('File Excel kosong!'));
                return false;
            }

            // Parse headers dan rows
            this.data.headers = data[0].map(h => String(h).trim());
            this.data.rows = data.slice(1)
                .filter(row => row.some(cell => cell && String(cell).trim() !== ""))
                .map(row => {
                    let obj = {};
                    this.data.headers.forEach((header, i) => {
                        obj[i] = (row[i] !== undefined && row[i] !== null) ? 
                            String(row[i]).trim() : "";
                    });
                    return obj;
                });

            // Auto-detect mappings
            this.autoDetectMappings();

            spinner.succeed(chalk.green(`Berhasil membaca: ${this.data.rows.length} baris, ${this.data.headers.length} kolom`));
            
            // Tampilkan preview
            console.log(chalk.cyan('\n📊 Preview Data:'));
            console.log(chalk.gray('Headers:'), this.data.headers.join(' | '));
            console.log(chalk.gray(`Total rows: ${this.data.rows.length}\n`));
            
            return true;
        } catch (e) {
            spinner.fail(chalk.red(`Error: ${e.message}`));
            return false;
        }
    }

    autoDetectMappings() {
        this.data.mappings = this.data.headers.map((header, index) => ({
            colIdx: index,
            header: header,
            type: this.guessFieldType(header),
            selector: `input[name="${header.toLowerCase().replace(/\s+/g, '_')}"]`
        }));
    }

    guessFieldType(header) {
        const lower = header.toLowerCase();
        if (lower.includes('kartu') || lower.includes('card') || lower.includes('nomor') || 
            lower.includes('number') || lower.includes('cc')) return 'card_number';
        if (lower.includes('expiry') || lower.includes('exp') || lower.includes('tanggal') || 
            lower.includes('date') || lower.includes('valid')) return 'expiry';
        if (lower.includes('cvv') || lower.includes('cvc') || lower.includes('cvv2') || 
            lower.includes('security') || lower.includes('kode')) return 'cvv';
        return 'text';
    }

    // ========== BROWSER AUTOMATION ==========
    async initBrowser() {
        const spinner = ora('Memulai browser...').start();
        try {
            this.browser = await puppeteer.launch({
                headless: CONFIG.headless,
                args: [
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-gpu'
                ]
            });
            this.page = await this.browser.newPage();
            await this.page.setViewport(CONFIG.viewport);
            spinner.succeed(chalk.green('Browser siap'));
        } catch (e) {
            spinner.fail(chalk.red(`Gagal memulai browser: ${e.message}`));
            throw e;
        }
    }

    async navigateToUrl(url) {
        const spinner = ora(`Navigasi ke ${url}...`).start();
        try {
            await this.page.goto(url, { 
                waitUntil: 'networkidle2',
                timeout: CONFIG.timeout 
            });
            spinner.succeed(chalk.green('Halaman berhasil dimuat'));
            return true;
        } catch (e) {
            spinner.fail(chalk.red(`Gagal navigasi: ${e.message}`));
            return false;
        }
    }

    async fillCheckoutForm(rowData, rowIndex) {
        console.log(chalk.cyan(`\n📝 Memproses baris ${rowIndex + 1}...`));
        
        try {
            // Extract data
            const cardNumber = rowData[0] || '';
            const expiry = rowData[1] || '';
            const cvv = rowData[2] || '';

            console.log(chalk.gray(`  Kartu: ${this.maskCard(cardNumber)}`));
            console.log(chalk.gray(`  Expiry: ${expiry || 'N/A'}`));
            console.log(chalk.gray(`  CVV: ${cvv ? '***' : 'N/A'}`));

            // Find and fill card number field
            const cardSelectors = [
                'input[name*="card"]',
                'input[name*="number"]',
                'input[name*="cc"]',
                'input[autocomplete="cc-number"]',
                '#cardNumber',
                '#cc-number'
            ];

            for (const selector of cardSelectors) {
                const field = await this.page.$(selector);
                if (field) {
                    await field.click({ clickCount: 3 });
                    await field.type(cardNumber, { delay: 50 });
                    console.log(chalk.green('  ✓ Nomor kartu terisi'));
                    break;
                }
            }

            // Find and fill expiry field
            const expirySelectors = [
                'input[name*="expiry"]',
                'input[name*="exp"]',
                'input[name*="date"]',
                'input[autocomplete="cc-exp"]',
                '#expiry',
                '#exp-date'
            ];

            for (const selector of expirySelectors) {
                const field = await this.page.$(selector);
                if (field) {
                    await field.click({ clickCount: 3 });
                    await field.type(expiry, { delay: 50 });
                    console.log(chalk.green('  ✓ Expiry terisi'));
                    break;
                }
            }

            // Find and fill CVV field
            const cvvSelectors = [
                'input[name*="cvv"]',
                'input[name*="cvc"]',
                'input[name*="security"]',
                'input[autocomplete="cc-csc"]',
                '#cvv',
                '#cvc'
            ];

            for (const selector of cvvSelectors) {
                const field = await this.page.$(selector);
                if (field) {
                    await field.click({ clickCount: 3 });
                    await field.type(cvv, { delay: 50 });
                    console.log(chalk.green('  ✓ CVV terisi'));
                    break;
                }
            }

            return true;
        } catch (e) {
            console.log(chalk.red(`  ✗ Error: ${e.message}`));
            return false;
        }
    }

    async clickPayButton() {
        console.log(chalk.yellow('  💳 Mengklik tombol bayar...'));
        
        const paySelectors = [
            'button[type="submit"]',
            'button:contains("Bayar")',
            'button:contains("Pay")',
            'input[type="submit"]',
            '#submit',
            '.btn-pay',
            '.checkout-btn'
        ];

        for (const selector of paySelectors) {
            try {
                const button = await this.page.$(selector);
                if (button) {
                    await button.click();
                    await this.page.waitForTimeout(2000);
                    console.log(chalk.green('  ✓ Tombol bayar diklik'));
                    return true;
                }
            } catch (e) {
                continue;
            }
        }

        console.log(chalk.yellow('  ⚠ Tombol bayar tidak ditemukan'));
        return false;
    }

    async handleOTP(rowIndex) {
        console.log(chalk.cyan('  🔐 Menangani OTP...'));
        
        // Wait for OTP page to load
        await this.page.waitForTimeout(2000);
        
        // Check if OTP input exists
        const otpSelectors = [
            'input[name*="otp"]',
            'input[type="tel"][maxlength="6"]',
            '#otp',
            '.otp-input'
        ];

        for (const selector of otpSelectors) {
            const otpField = await this.page.$(selector);
            if (otpField) {
                // Simulate OTP input (dummy OTP: 123456)
                const dummyOTP = '123456';
                await otpField.click();
                await otpField.type(dummyOTP, { delay: 100 });
                console.log(chalk.green('  ✓ OTP terisi: 123456'));
                
                // Click verify button
                const verifySelectors = [
                    'button:contains("Verifikasi")',
                    'button:contains("Verify")',
                    'button[type="submit"]',
                    '#verify',
                    '.otp-btn'
                ];

                for (const vSelector of verifySelectors) {
                    const verifyBtn = await this.page.$(vSelector);
                    if (verifyBtn) {
                        await verifyBtn.click();
                        await this.page.waitForTimeout(3000);
                        console.log(chalk.green('  ✓ OTP diverifikasi'));
                        
                        // Save to history
                        this.addToHistory({
                            status: 'success',
                            row: rowIndex,
                            card: this.maskCard(this.data.rows[rowIndex][0] || ''),
                            timestamp: new Date().toISOString()
                        });
                        
                        return true;
                    }
                }
                break;
            }
        }

        console.log(chalk.yellow('  ⚠ Halaman OTP tidak terdeteksi'));
        return false;
    }

    async handleError(rowIndex) {
        console.log(chalk.red('  ❌ Kartu ditolak!'));
        
        // Save to history
        this.addToHistory({
            status: 'failed',
            row: rowIndex,
            card: this.maskCard(this.data.rows[rowIndex][0] || ''),
            timestamp: new Date().toISOString()
        });

        // Wait for auto-reset or manual reset
        await this.page.waitForTimeout(CONFIG.errorResetDelay);
        console.log(chalk.yellow('  🔄 Auto-reset...'));
        
        return false;
    }

    addToHistory(entry) {
        this.data.history.unshift({
            id: Date.now(),
            ...entry
        });
        this.saveHistory();
    }

    maskCard(cardNumber) {
        if (!cardNumber) return '----';
        const cleaned = cardNumber.replace(/\D/g, '');
        if (cleaned.length >= 8) {
            return `${cleaned.slice(0, 4)} **** **** ${cleaned.slice(-4)}`;
        }
        return cardNumber;
    }

    // ========== MAIN AUTOMATION ==========
    async runAutomation(url, mode = 'payment') {
        console.log(chalk.blue.bold('\n🚀 Memulai Otomasi FormFlow'));
        console.log(chalk.gray(`Mode: ${mode}`));
        console.log(chalk.gray(`URL: ${url}`));
        console.log(chalk.gray(`Total data: ${this.data.rows.length} baris\n`));

        if (!this.browser) {
            await this.initBrowser();
        }

        if (!await this.navigateToUrl(url)) {
            console.log(chalk.red('Gagal memuat halaman target'));
            return;
        }

        let successCount = 0;
        let failCount = 0;

        for (let i = 0; i < this.data.rows.length; i++) {
            const row = this.data.rows[i];
            
            console.log(chalk.blue(`\n${'='.repeat(50)}`));
            console.log(chalk.blue(`Baris ${i + 1}/${this.data.rows.length}`));
            console.log(chalk.blue(`${'='.repeat(50)}`));

            // Fill the form
            const fillSuccess = await this.fillCheckoutForm(row, i);
            
            if (fillSuccess) {
                await this.page.waitForTimeout(1000);
                
                // Click pay button
                await this.clickPayButton();
                
                // Check if OTP or error page
                const currentUrl = this.page.url();
                const pageContent = await this.page.content();
                
                if (pageContent.includes('otp') || pageContent.includes('OTP') || 
                    pageContent.includes('verifikasi') || pageContent.includes('verify')) {
                    // Handle OTP
                    const otpSuccess = await this.handleOTP(i);
                    if (otpSuccess) {
                        successCount++;
                        console.log(chalk.green(`✅ Baris ${i + 1}: SUKSES`));
                    } else {
                        failCount++;
                        console.log(chalk.yellow(`⚠ Baris ${i + 1}: OTP GAGAL`));
                    }
                } else if (pageContent.includes('error') || pageContent.includes('ditolak') || 
                          pageContent.includes('declined') || pageContent.includes('gagal')) {
                    // Handle error
                    await this.handleError(i);
                    failCount++;
                    console.log(chalk.red(`❌ Baris ${i + 1}: DITOLAK`));
                } else {
                    // Unknown state
                    console.log(chalk.yellow(`⚠ Baris ${i + 1}: Status tidak diketahui`));
                    failCount++;
                }
            } else {
                console.log(chalk.red(`❌ Baris ${i + 1}: Gagal mengisi form`));
                failCount++;
            }

            // Reset or navigate back for next card
            if (i < this.data.rows.length - 1) {
                console.log(chalk.gray('  🔄 Menyiapkan kartu berikutnya...'));
                await this.navigateToUrl(url);
            }
        }

        // Show summary
        console.log(chalk.blue(`\n${'='.repeat(50)}`));
        console.log(chalk.blue.bold('📊 RINGKASAN'));
        console.log(chalk.blue(`${'='.repeat(50)}`));
        console.log(chalk.green(`✅ Sukses: ${successCount}`));
        console.log(chalk.red(`❌ Gagal: ${failCount}`));
        console.log(chalk.gray(`📊 Total: ${this.data.rows.length}`));
        console.log(chalk.blue(`${'='.repeat(50)}\n`));
    }

    // ========== INTERACTIVE MODE ==========
    async showHistory() {
        console.log(chalk.blue.bold('\n📋 Riwayat Transaksi'));
        console.log(chalk.blue('='.repeat(60)));

        if (this.data.history.length === 0) {
            console.log(chalk.gray('Belum ada riwayat transaksi'));
            return;
        }

        this.data.history.forEach((entry, index) => {
            const status = entry.status === 'success' ? 
                chalk.green('✅ SUKSES') : chalk.red('❌ GAGAL');
            const date = new Date(entry.timestamp).toLocaleString('id-ID');
            
            console.log(chalk.cyan(`\n[${index + 1}] ${date}`));
            console.log(`  Status: ${status}`);
            console.log(`  Kartu: ${entry.card}`);
            console.log(`  Row: ${entry.row + 1}`);
        });
        
        console.log(chalk.blue('\n' + '='.repeat(60)));
    }

    async exportHistory(format = 'csv') {
        if (this.data.history.length === 0) {
            console.log(chalk.yellow('Tidak ada data untuk diexport'));
            return;
        }

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const filename = `formflow_history_${timestamp}.${format}`;

        try {
            if (format === 'csv') {
                const csv = this.convertToCSV(this.data.history);
                fs.writeFileSync(filename, csv);
            } else if (format === 'json') {
                fs.writeFileSync(filename, JSON.stringify(this.data.history, null, 2));
            } else if (format === 'xlsx') {
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.json_to_sheet(this.data.history);
                XLSX.utils.book_append_sheet(wb, ws, 'History');
                XLSX.writeFile(wb, filename);
            }

            console.log(chalk.green(`✅ History diexport ke: ${filename}`));
        } catch (e) {
            console.log(chalk.red(`❌ Gagal export: ${e.message}`));
        }
    }

    convertToCSV(data) {
        if (data.length === 0) return '';
        
        const headers = Object.keys(data[0]);
        const csvRows = [];
        
        // Header
        csvRows.push(headers.join(','));
        
        // Data rows
        for (const row of data) {
            const values = headers.map(header => {
                const val = row[header] || '';
                return `"${String(val).replace(/"/g, '""')}"`;
            });
            csvRows.push(values.join(','));
        }
        
        return csvRows.join('\n');
    }

    // ========== CLEANUP ==========
    async close() {
        if (this.browser) {
            await this.browser.close();
            console.log(chalk.gray('\nBrowser ditutup'));
        }
    }
}

// ========== CLI INTERFACE ==========
async function main() {
    console.log(chalk.blue.bold('\n╔══════════════════════════════════╗'));
    console.log(chalk.blue.bold('║      FormFlow CLI v1.0          ║'));
    console.log(chalk.blue.bold('║   Auto Form Filler for Linux    ║'));
    console.log(chalk.blue.bold('╚══════════════════════════════════╝\n'));

    const formflow = new FormFlow();

    // Parse command line arguments
    const args = process.argv.slice(2);
    let command = args[0];
    let option = args[1];

    if (!command) {
        // Interactive mode
        command = readline.keyInSelect(
            ['Jalankan Otomasi', 'Lihat Riwayat', 'Export Riwayat', 'Keluar'],
            'Pilih menu: '
        );

        switch (command) {
            case 0: // Run automation
                const filePath = readline.questionPath('📂 Path file Excel: ', {
                    exists: true,
                    extensions: ['xlsx', 'xls', 'csv']
                });

                const url = readline.question('🌐 URL target: ', {
                    defaultInput: 'https://example.com/checkout'
                });

                const mode = readline.keyInSelect(
                    ['Pembayaran (Payment)', 'Login', 'Registrasi'],
                    'Pilih mode: '
                );

                const modes = ['payment', 'login', 'registration'];
                
                if (await formflow.loadExcelFile(filePath)) {
                    await formflow.runAutomation(url, modes[mode] || 'payment');
                }
                break;

            case 1: // View history
                await formflow.showHistory();
                break;

            case 2: // Export history
                const format = readline.keyInSelect(
                    ['CSV', 'JSON', 'Excel'],
                    'Pilih format export: '
                );
                const formats = ['csv', 'json', 'xlsx'];
                await formflow.exportHistory(formats[format] || 'csv');
                break;

            case 3: // Exit
                console.log(chalk.gray('Sampai jumpa! 👋'));
                process.exit(0);
        }
    } else {
        // Command line mode
        switch (command) {
            case 'run':
                if (!option || !args[2]) {
                    console.log(chalk.red('Usage: formflow run <file.xlsx> <url> [mode]'));
                    process.exit(1);
                }
                const mode = args[3] || 'payment';
                if (await formflow.loadExcelFile(option)) {
                    await formflow.runAutomation(args[2], mode);
                }
                break;

            case 'history':
                await formflow.showHistory();
                break;

            case 'export':
                const fmt = option || 'csv';
                await formflow.exportHistory(fmt);
                break;

            case 'help':
                console.log(chalk.cyan('\nFormFlow CLI Usage:'));
                console.log('  formflow run <file.xlsx> <url> [mode]  - Jalankan otomasi');
                console.log('  formflow history                        - Lihat riwayat');
                console.log('  formflow export [csv|json|xlsx]        - Export riwayat');
                console.log('  formflow                                - Mode interaktif\n');
                break;

            default:
                console.log(chalk.red(`Command tidak dikenal: ${command}`));
                console.log(chalk.gray('Gunakan "formflow help" untuk bantuan'));
                process.exit(1);
        }
    }

    // Cleanup
    await formflow.close();
}

// Handle CTRL+C
process.on('SIGINT', async () => {
    console.log(chalk.yellow('\n\n⚠ Program dihentikan'));
    process.exit(0);
});

// Run main
main().catch(e => {
    console.error(chalk.red('\n❌ Error:'), e.message);
    process.exit(1);
});