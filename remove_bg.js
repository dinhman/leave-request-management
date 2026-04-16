const fs = require('fs');
const PNG = require('pngjs').PNG;
const path = require('path');

const fileColor = path.join(__dirname, 'teams', 'f444d51b-0e1a-44c4-af9e-084a4f7fce0f_color.png');
const fileOutline = path.join(__dirname, 'teams', 'f444d51b-0e1a-44c4-af9e-084a4f7fce0f_outline.png');

function processImage(filePath, isOutline) {
    if (!fs.existsSync(filePath)) {
        console.log(`File not found: ${filePath}`);
        return;
    }
    fs.createReadStream(filePath)
        .pipe(new PNG({ filterType: 4 }))
        .on('parsed', function () {
            for (let y = 0; y < this.height; y++) {
                for (let x = 0; x < this.width; x++) {
                    const idx = (this.width * y + x) << 2;
                    const r = this.data[idx];
                    const g = this.data[idx + 1];
                    const b = this.data[idx + 2];
                    
                    // If the pixel is close to white, make it fully transparent
                    if (r > 230 && g > 230 && b > 230) {
                        this.data[idx + 3] = 0; // Alpha to 0
                    } else if (isOutline) {
                        // For outline image, if it's not white, it MUST be white and opaque
                        this.data[idx] = 255;
                        this.data[idx + 1] = 255;
                        this.data[idx + 2] = 255;
                        this.data[idx + 3] = 255;
                    }
                }
            }
            this.pack().pipe(fs.createWriteStream(filePath)).on('finish', () => {
                console.log(`Processed ${filePath}`);
            });
        });
}

processImage(fileColor, false);
processImage(fileOutline, true);
