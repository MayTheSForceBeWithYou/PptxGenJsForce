import { LightningElement } from 'lwc';
import { loadScript } from 'lightning/platformResourceLoader';
import PPTXGEN from '@salesforce/resourceUrl/pptxGen';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';

export default class TestPptxGenDownload extends LightningElement {
    isPptxGenLoaded = false;
    
    connectedCallback() {
        if(!this.isPptxGenLoaded) {
            Promise.all([
                loadScript(this, PPTXGEN)
            ]).then(() => {
                console.log('pptxGen static resource successfully loaded');
                this.isPptxGenLoaded = true;
            }).catch((err) => {
                console.error('Error loading pptxGen static resource');
                console.error(JSON.stringify(err, null, 2));
                this.dispatchEvent(
                    new ShowToastEvent({
                        title: 'Error loading pptxGen',
                        message: JSON.stringify(err, null, 2),
                        variant: 'error',
                    }),
                );
            });
        }
    }
    
    renderedCallback() {
        if(this.isPptxGenLoaded) {
            console.log('Instantiating pptx');
            if(window.PptxGenJS) {
                let pptx = new window.PptxGenJS();
                console.log('Adding slide');
                let slide01 = pptx.addSlide();
                slide01.background = { color: "#456ff6" };
                
                console.log('Adding text to slide01');
                slide01.addText("Sales Report", {
                    x: 0,
                    y: 1,
                    w: "50%",
                    h: 1,
                    align: "center",
                    color: "#eff0f1",
                    fill: "#456ff6",
                    fontSize: 52,
                });
                const monthYearText = new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long' });
                slide01.addText(monthYearText, {
                    x: 0,
                    y: 2,
                    w: "50%",
                    h: 1,
                    align: "center",
                    color: "#eff0f1",
                    fill: "#456ff6",
                    fontSize: 28,
                });
                
                console.log('Writing file');
                pptx.write("base64")
                .then((data) => {
                    console.log('write as base64: Here are 0-100 chars of `data`:\n');
                    console.log(data.substring(0, 100));
                    const aHref = 'data:application/vnd.ms-powerpoint;base64,' + data;
                    let downloadElement = document.createElement('a');
                    downloadElement.href = aHref;
                    downloadElement.target = '_self';
                    downloadElement.download = 'Demo PptxGen.pptx';
                    const downloadDiv = this.querySelector('.download-div');
                    document.body.appendChild(downloadElement);
                    downloadElement.click();
                })
                .catch(err => {
                    this.dispatchEvent(
                        new ShowToastEvent({
                            title: 'Error creating pptx file',
                            message: JSON.stringify(err, null, 2),
                            variant: 'error',
                        }),
                    );
                });
            } else {
                console.log('window.PptxGenJS is falsy');
            }
        }
    }
}