const QRCode = require('qrcode');
const fs = require('fs');
const path = require('path');

// Generate QR code for the investor pitch webpage
// URL will be updated once deployed
const PLACEHOLDER_URL = 'https://investor-pitch.onestrategy.app';

async function generateQR() {
    const outputPath = path.join(__dirname, 'assets', 'qr-code.png');
    
    // Ensure assets directory exists
    if (!fs.existsSync(path.join(__dirname, 'assets'))) {
        fs.mkdirSync(path.join(__dirname, 'assets'), { recursive: true });
    }

    // Generate QR code
    await QRCode.toFile(outputPath, PLACEHOLDER_URL, {
        type: 'png',
        width: 400,
        margin: 2,
        color: {
            dark: '#D86DCB',  // Pink accent color
            light: '#0a0a0f'  // Dark background
        }
    });

    console.log('QR Code generated:', outputPath);

    // Also generate white background version for PDF
    const outputPathWhite = path.join(__dirname, 'assets', 'qr-code-white.png');
    await QRCode.toFile(outputPathWhite, PLACEHOLDER_URL, {
        type: 'png',
        width: 400,
        margin: 2,
        color: {
            dark: '#000000',
            light: '#FFFFFF'
        }
    });

    console.log('QR Code (white bg) generated:', outputPathWhite);

    // Generate SVG version
    const svgPath = path.join(__dirname, 'assets', 'qr-code.svg');
    const svgString = await QRCode.toString(PLACEHOLDER_URL, {
        type: 'svg',
        width: 200,
        margin: 1,
        color: {
            dark: '#D86DCB',
            light: '#0a0a0f'
        }
    });
    fs.writeFileSync(svgPath, svgString);
    console.log('QR Code SVG generated:', svgPath);
}

generateQR().catch(console.error);
