// Powerbi
import powerbiVisualsApi from "powerbi-visuals-api";

// Chartjs
import Chart from 'chart.js/auto';
import ChartDataLabels from 'chartjs-plugin-datalabels';

// Power BI Visuals imports
import IVisual = powerbiVisualsApi.extensibility.IVisual;
import VisualConstructorOptions = powerbiVisualsApi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbiVisualsApi.extensibility.visual.VisualUpdateOptions;
import DataView = powerbiVisualsApi.DataView;

// Selection manager
import ISelectionManager = powerbi.extensibility.ISelectionManager;

// Rendering events
import IVisualEventService = powerbi.extensibility.IVisualEventService;

// Chart.js custom visual
export class PolarAreaChart implements IVisual {
    private chart: Chart<any, any, any>;
    private host: powerbiVisualsApi.extensibility.visual.IVisualHost;
    private canvas: HTMLCanvasElement;
    private categoryName = "";
    private selectionManager: ISelectionManager;
    private events: IVisualEventService;
    private element: HTMLElement;

    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    //// CHART PLUGINS ////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    // The plugin draws the following
    // - Color full (i.e., most) outer circle with the background color(s) from the format pane
    // - Color the outer circle (i.e., inside full outer) with alpha 1 (i.e., no opacity)
    // - Draw type labels
    // - Draw the full (i.e., most) outer circle
    // - Draw the outer circle (i.e., inside full outer)
    // - Draw lines between segments - outer
    // - Draw lines between segments - inner
    // - Draw category names

    // We use helper functions before the actual drawing in the plugin:
    // 0. Set default values for formatting pane
    // 1. Draw Type labels (helper function)
    // 2. Draw Angle Segments (helper function)
    // 3. Calculate the font size (helper functions)
    // 4. Draw category names and type names (helper functions)
    // 5. Create the plugins


    // 0. SET DEFAULT VALUES FOR FORMATTING PANE
    private bgColorType1: string = '#fff';
    private bgColorType2: string = '#fff';
    private bgColorType3: string = '#fff';
    private bgColorType4: string = '#fff';
    private bgColorType5: string = '#fff';
    private bgColorType6: string = '#fff';
    private bgColorType7: string = '#fff';
    private bgColorType8: string = '#fff';
    private bgColorType9: string = '#fff';
    private bgColorType10: string = '#fff';
    private bgColorCategory1: string = '#d3d3d3';
    private bgColorCategory2: string = '#d3d3d3';
    private bgColorCategory3: string = '#d3d3d3';
    private bgColorCategory4: string = '#d3d3d3';
    private bgColorCategory5: string = '#d3d3d3';
    private bgColorCategory6: string = '#d3d3d3';
    private bgColorCategory7: string = '#d3d3d3';
    private bgColorCategory8: string = '#d3d3d3';
    private bgColorCategory9: string = '#d3d3d3';
    private bgColorCategory10: string = '#d3d3d3';
    private bgColorCategory11: string = '#d3d3d3';
    private bgColorCategory12: string = '#d3d3d3';
    private bgColorCategory13: string = '#d3d3d3';
    private bgColorCategory14: string = '#d3d3d3';
    private bgColorCategory15: string = '#d3d3d3';
    private bgColorCategory16: string = '#d3d3d3';
    private bgColorCategory17: string = '#d3d3d3';
    private bgColorCategory18: string = '#d3d3d3';
    private bgColorCategory19: string = '#d3d3d3';
    private bgColorCategory20: string = '#d3d3d3';
    private bgColorCategory21: string = '#d3d3d3';
    private bgColorCategory22: string = '#d3d3d3';
    private bgColorCategory23: string = '#d3d3d3';
    private bgColorCategory24: string = '#d3d3d3';
    private bgColorCategory25: string = '#d3d3d3';
    private bgColorCategory26: string = '#d3d3d3';
    private bgColorCategory27: string = '#d3d3d3';
    private bgColorCategory28: string = '#d3d3d3';
    private bgColorCategory29: string = '#d3d3d3';
    private bgColorCategory30: string = '#d3d3d3';
    private fontOuterCircle: string = 'arial';

    // Padding calculation
    private calculatePercentagePadding(chart) {
        
        const chartArea = chart.chartArea;
        const innerRadius = Math.min(chartArea.right - chartArea.left, chartArea.bottom - chartArea.top) / 2; 
        const outerRadius = innerRadius + Math.min(chartArea.right - chartArea.left, chartArea.bottom - chartArea.top) / 14 * 1;
        const fullOuterRadius = outerRadius + Math.min(chartArea.right - chartArea.left, chartArea.bottom - chartArea.top) / 14 * 1;

        // Calculate padding based on the outer circle's radius
        // Ensuring the outer circle fits within the canvas
        const requiredPadding = (fullOuterRadius-innerRadius)*1.2; 

        return requiredPadding;
    }

    // 1. DRAW TYPE LABELS (HELPER FUNCTION)
    private drawTypeLabel(ctx, chartArea, angle, innerRadius, outerRadius, label) {
        const centerX = chartArea.left + (chartArea.right - chartArea.left) / 2;
        const centerY = chartArea.top + (chartArea.bottom - chartArea.top) / 2;
        const labelRadius = (innerRadius + outerRadius) / 2;
        const labelText = label.trimEnd();

        // Calculate position for each letter
        for (let i = 0; i < labelText.length; i++) {
            const charAngle = angle + (i - labelText.length / 2) * 0.035;

            // Rotate text by 180 degrees if it's in the bottom half of the circle
            let rotationAngle = charAngle;
            if (angle > 0 && angle < Math.PI) {
                rotationAngle +=  + Math.PI; // Add 180 degrees
            }

            // Continue
            const fontSize = labelRadius / 25; // Adjust font size based on radius
            const labelX = centerX + labelRadius * Math.cos(charAngle);
            const labelY = centerY + labelRadius * Math.sin(charAngle);
            ctx.save();
            ctx.translate(labelX, labelY);
            ctx.rotate(rotationAngle + Math.PI / 2);
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.font = `${fontSize}px ${this.fontOuterCircle}`;
            ctx.fillStyle = 'black';
            if (angle > 0 && angle < Math.PI) { // If lower half, get last letter first and vice versa
                ctx.fillText(labelText[labelText.length-1-i], 0, 0);
            }else{
                ctx.fillText(labelText[i], 0, 0);
            }
            ctx.restore();
        }
    }


    // 2. DRAW ANGLE SEGMENTS (HELPER FUNCTION)
    private drawAngleLine(ctx, chartArea, angle, innerRadius, outerRadius) {
        const centerX = chartArea.left + (chartArea.right - chartArea.left) / 2;
        const centerY = chartArea.top + (chartArea.bottom - chartArea.top) / 2;
        const startX = centerX + innerRadius * Math.cos(angle);
        const startY = centerY + innerRadius * Math.sin(angle);
        const endX = centerX + outerRadius * Math.cos(angle);
        const endY = centerY + outerRadius * Math.sin(angle);
        ctx.beginPath();
        ctx.moveTo(startX, startY);
        ctx.lineTo(endX, endY);
        ctx.stroke();
    }


    // 3. CALCULATE THE FONT SIZE (HELPER FUNCTIONS)
    private calculateAvailableSpace(innerRadius, outerRadius) {
        // Average radius for the text
        const avgRadius = (innerRadius + outerRadius) / 8;
        // Use a fraction of arc length as available space to ensure text fits comfortably
        return avgRadius * 0.9; 
    }

    private adjustFontSizeForSpace(ctx, text, availableSpace, outerRadius) {
        const maxFontSize = outerRadius / 25;
        const minFontSize = outerRadius / 40;
        let fontSize = maxFontSize;
        do {
            ctx.font = `${fontSize}px ${this.fontOuterCircle}`;
            if (ctx.measureText(text).width < availableSpace) {
                return fontSize; 
            }
            fontSize--;
        } while (fontSize > minFontSize);
    
        return fontSize; 
    }


    // 4. DRAW CATEGORY NAMES AND TYPE NAMES (HELPER FUNCTIONS)
    private drawCategoryNames(chart,segmentAngle,centerX,centerY,innerRadius,outerRadius,ctx){
        chart.data.labels.forEach((label, index) => {
            const angle = segmentAngle * index + segmentAngle / 2 - Math.PI / 2; // Adjusted for centering within segment
            const labelRadius = (innerRadius + outerRadius) / 2;
            const labelX = centerX + labelRadius * Math.cos(angle);
            const labelY = centerY + labelRadius * Math.sin(angle);

            // Calculate available space for text
            const availableSpace = this.calculateAvailableSpace(innerRadius, outerRadius);           
            ctx.save();
            ctx.translate(labelX, labelY);
            let rotationAngle = angle + Math.PI / 2;
            if (angle > 0 && angle < Math.PI) {
                rotationAngle += Math.PI; // Add 180 degrees
            }
            ctx.rotate(rotationAngle);

            // Adjust font size based on available space
            const fontSize = this.adjustFontSizeForSpace(ctx, label, availableSpace, outerRadius);
            ctx.font = `${fontSize}px ${this.fontOuterCircle}`;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillStyle = 'black';
            
            // Split label into words
            const words = label.split(' ');
            if (words.length > 2) {
                const middleIndex = Math.ceil(words.length / 2);
                const topLine = words.slice(0, middleIndex).join(' ');
                const bottomLine = words.slice(middleIndex).join(' ');
                ctx.fillText(topLine, 0, -fontSize / 2);
                ctx.fillText(bottomLine, 0, fontSize / 2);
            } else if (words.length === 2) {
                ctx.fillText(words[0], 0, -fontSize / 2);
                ctx.fillText(words[1], 0, fontSize / 2);
            } else { // Shorten word if it's too long
                const maxWordLength = fontSize * 10;
                let word = words[0];
                while (ctx.measureText(word).width > maxWordLength) {
                    word = word.substring(0, word.length - 1);
                }
                ctx.fillText(word, 0, 0);
            }
            ctx.restore();
        });
    }

    private drawTypesNames(chart,ctx,chartArea,outerRadius,fullOuterRadius){
        // Draw type labels
        let previousType = null;
        const typeAngles = new Map();
        const totalSegments = chart.data.labels.length;
        const segmentAngleType = 2 * Math.PI / totalSegments;
        chart.data.datasets.forEach((dataset) => {
            dataset.data.forEach((value, index) => {
                const currentType = dataset.dataType[index];
                const angle = segmentAngleType * index - Math.PI / 2;
                const currentTypeIndex = `${currentType}${' '.repeat(index)}`;
                let hasCurrentType = 0;
                if (currentType !== previousType) {  // !typeAngles.has(currentType)
                    if(!typeAngles.has(currentType)){
                        typeAngles.set(currentType, { start: angle, count: 0 });
                        hasCurrentType = 0;
                    }else{
                        typeAngles.set(currentTypeIndex, { start: angle, count: 0 });
                        hasCurrentType = 1;
                    }
                }
                if(hasCurrentType === 0){
                    typeAngles.get(currentType).end = segmentAngleType * (index+1) - Math.PI / 2;
                    typeAngles.get(currentType).count++;
                }else{
                    typeAngles.get(currentTypeIndex).end = segmentAngleType * (index+1) - Math.PI / 2;
                    typeAngles.get(currentTypeIndex).count++;
                }
                if (currentType !== previousType) {
                    this.drawAngleLine(ctx, chartArea, angle, outerRadius, fullOuterRadius);
                }
                previousType = currentType;
            });
        });
        typeAngles.forEach((angles, type) => {
            const middleAngle = (angles.start + angles.end) / 2;
            this.drawTypeLabel(ctx, chartArea, middleAngle, outerRadius, fullOuterRadius, type);
        });
    }

    // 5. CREATE THE PLUGINS
    private drawOuterCirclePlugin(chart) {
        if(chart){
            const ctx = chart.ctx;
            const chartArea = chart.chartArea;
            const centerX = chartArea.left + (chartArea.right - chartArea.left) / 2;
            const centerY = chartArea.top + (chartArea.bottom - chartArea.top) / 2;
            const innerRadius = Math.min(chartArea.right - chartArea.left, chartArea.bottom - chartArea.top) / 2; 
            const outerRadius = innerRadius + Math.min(chartArea.right - chartArea.left, chartArea.bottom - chartArea.top) / 14 * 1;
            const fullOuterRadius = outerRadius + Math.min(chartArea.right - chartArea.left, chartArea.bottom - chartArea.top) / 14 * 1;
            const segmentAngleFull = 2 * Math.PI / chart.data.labels.length;

            // Color full (i.e., most) outer circle with the background color(s) from the format pane
            chart.data.datasets.forEach((dataset) => {
                for (let i = 0; i < chart.data.labels.length; i++) {
                    const startAngle = segmentAngleFull * i - Math.PI / 2;
                    const endAngle = startAngle + segmentAngleFull;
                    let color = dataset.backgroundColorType[i % dataset.backgroundColorType.length];
                    if(color === null){color = this.hexToRgba('#d3d3d3',1);}
                    ctx.beginPath();
                    ctx.arc(centerX, centerY, outerRadius, startAngle, endAngle);
                    ctx.arc(centerX, centerY, fullOuterRadius, endAngle, startAngle, true);
                    ctx.closePath();
                    ctx.fillStyle = color; 
                    ctx.fill();
                }
            });

            // Color the outer circle (i.e., inside full outer) with alpha 1 (i.e., no opacity)
            chart.data.datasets.forEach((dataset) => {
                for (let i = 0; i < chart.data.labels.length; i++) {
                    const startAngle = segmentAngleFull * i - Math.PI / 2;
                    const endAngle = startAngle + segmentAngleFull;
                    let color = this.hexToRgba('#d3d3d3',0.5);
                    if(dataset.backgroundColor[i % dataset.backgroundColor.length] !== null){
                        try { color = this.adjustRgbaAlpha(dataset.backgroundColor[i % dataset.backgroundColor.length], 1); } catch (error) { color = dataset.backgroundColor[i % dataset.backgroundColor.length]; }
                    }
                    ctx.beginPath();
                    ctx.arc(centerX, centerY, innerRadius, startAngle, endAngle);
                    ctx.arc(centerX, centerY, outerRadius, endAngle, startAngle, true);
                    ctx.closePath();
                    ctx.fillStyle = color;
                    ctx.fill();
                }
            });

            // Draw the full (i.e., most) outer circle
            ctx.beginPath();
            ctx.arc(centerX, centerY, fullOuterRadius, 0, 2 * Math.PI);
            ctx.lineWidth = 1.5;
            ctx.strokeStyle = 'black';
            ctx.stroke();

            // Draw lines between segments - outer
            const segmentAngle = 2 * Math.PI / chart.data.labels.length;
            for (let i = 0; i < chart.data.labels.length; i++) {
                const angle = segmentAngle * i - Math.PI / 2;
                const startX = centerX + innerRadius * Math.cos(angle);
                const startY = centerY + innerRadius * Math.sin(angle);
                const endX = centerX + outerRadius * Math.cos(angle);
                const endY = centerY + outerRadius * Math.sin(angle);
                ctx.beginPath();
                ctx.moveTo(startX, startY);
                ctx.lineTo(endX, endY);
                ctx.stroke();
            }

            // Draw the outer circle (i.e., inside full outer)
            ctx.save();
            [outerRadius, innerRadius].forEach(radius => {
                ctx.beginPath();
                ctx.arc(centerX, centerY, radius, 0, 2 * Math.PI);
                ctx.lineWidth = 1.5;
                ctx.strokeStyle = 'black';
                ctx.stroke();
            });

            // Draw lines between segments - inner
            for (let i = 0; i < chart.data.labels.length; i++) {
                const angle = segmentAngle * i - Math.PI / 2;
                const startX = centerX;
                const startY = centerY;
                const endX = centerX + innerRadius * Math.cos(angle);
                const endY = centerY + innerRadius * Math.sin(angle);
                ctx.beginPath();
                ctx.lineWidth = 1.5;
                ctx.strokeStyle = 'black';
                ctx.moveTo(startX, startY);
                ctx.lineTo(endX, endY);
                ctx.stroke();
            }

            // Draw names
            this.drawCategoryNames(chart,segmentAngle,centerX,centerY,innerRadius,outerRadius,ctx);
            if(chart.data.datasets[0]){
                if(chart.data.datasets[0].dataType[0] !== null){
                    this.drawTypesNames(chart,ctx,chartArea,outerRadius,fullOuterRadius);
                }
            }

            // Draw second measure lines
            try { this.drawSecondMeasureLines(chart, ctx, chart.data.datasets[0].secondValues); } catch (error) { 
                console.log("No second measure")
                // No need for more error handling - the visual can function without second measure, we will not show the extra line then
            }
            ctx.restore();
        }        
    } 

    // This function is called by the plugin to draw additional elements like the second measure line
    private drawSecondMeasureLines(chart, ctx, secondValues) {
        const data = chart.data;
        if (!data.datasets.length) return;
    
        const chartArea = chart.chartArea;
        const centerX = chartArea.left + (chartArea.right - chartArea.left) / 2;
        const centerY = chartArea.top + (chartArea.bottom - chartArea.top) / 2;
        const maxRadius = Math.min(chartArea.right - chartArea.left, chartArea.bottom - chartArea.top) / 2;
        const minRadius = chart.scales.r.min; 
        const maxRadiusValue = chart.scales.r.max; 
    
        data.labels.forEach((label, index, labels) => {
            const secondValue = secondValues[index];
            const valueRatio = secondValue / maxRadiusValue;
            const segmentRadius = minRadius + (maxRadius - minRadius) * valueRatio;
    
            // Calculate start and end angles for the arc
            const startAngle = (2 * Math.PI / labels.length) * index - Math.PI / 2; // Adjust the angle based on the index
            const endAngle = (2 * Math.PI / labels.length) * (index + 1) - Math.PI / 2; // Next segment
    
            // Draw the arc
            ctx.beginPath();
            ctx.arc(centerX, centerY, segmentRadius, startAngle, endAngle);
            ctx.strokeStyle = 'black';
            ctx.lineWidth = 1.5; 
            ctx.stroke();
        });
    }
    
    
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    //// CONSTRUCTOR ////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    private colored = -1;  // Start colored yes/no for filter out
    private categories_org_order = [];
    private previous_values = [];

    constructor(options: VisualConstructorOptions) {
        this.events = options.host.eventService;
        this.element = options.element;
        this.previous_values = []; 
        this.categories_org_order = [];
        this.host = options.host;
        this.canvas = document.createElement('canvas');
        options.element.appendChild(this.canvas);

        // On click > Filter out
        this.canvas.addEventListener('click', (event) => {
            
            // Initialize the selectionManager if it's not already done
            if (!this.selectionManager) {
                this.selectionManager = this.host.createSelectionManager();
            }
            
            // Get the elements at the clicked position
            const elements = this.chart.getElementsAtEventForMode(event, 'nearest', { intersect: true }, false);
            if (elements.length > 0) {
                const firstElement = elements[0];
                const datasetIndex = firstElement.datasetIndex;
                const dataIndex = firstElement.index;
                const dataPoint = this.chart.data.datasets[datasetIndex].selectionId[dataIndex];
                
                // Apply filter to other visuals based on the selection
                if (dataPoint && dataPoint.selection) {
                    this.selectionManager.select(dataPoint.selection).then(() => {
                        this.chart.data.datasets[0].backgroundColor.forEach((color, index) => {
                            if (index === dataIndex) {
                                if (dataIndex !== this.colored) {
                                    this.chart.data.datasets[datasetIndex].backgroundColor[dataIndex] = this.adjustRgbaAlpha(this.chart.data.datasets[datasetIndex].backgroundColor[dataIndex],1);
                                    this.colored = dataIndex;
                                }else{
                                    this.chart.data.datasets[datasetIndex].backgroundColor[dataIndex] = this.adjustRgbaAlpha(this.chart.data.datasets[datasetIndex].backgroundColor[dataIndex],0.5);
                                    this.colored = -1;
                                }
                            }else{
                                this.chart.data.datasets[0].backgroundColor[index] = this.adjustRgbaAlpha(this.chart.data.datasets[0].backgroundColor[index],0.5);
                            }
                        });
                    }).catch((error) => { console.error('Error applying filter:', error); });
                }else{
                    this.colored = -1;
                }
            } else {
                this.selectionManager.clear().then(() => { // Click outside element > clear
                    try {
                        this.chart.data.datasets[0].backgroundColor.forEach((color, index) => {
                            this.chart.data.datasets[0].backgroundColor[index] = this.adjustRgbaAlpha(this.chart.data.datasets[0].backgroundColor[index],0.5);
                        });
                    } catch (error) { 
                        console.log('no bg colors yet');
                        // No need for more error handling - the visual can function without the background colors, we just keep the default in that case
                    }
                }).catch((error) => { console.error('Error clearing selection:', error); });
                this.colored = -1;
            }
        });

        this.contextMenuStart();

        this.chart = new Chart(this.canvas.getContext('2d'), {
            type: 'polarArea',
            data: null,
            options: {
                devicePixelRatio: 3,
                responsive: true,
                maintainAspectRatio: false,
                layout: { padding: 100 },
                scales: {
                    r: { min: 0, max: 100, ticks: { stepSize: 20, backdropColor: 'rgba(0, 0, 0, 0)' } }
                },
                plugins: {
                    legend: { display: false, },
                    tooltip: { enabled: true, },
                    datalabels: {
                        anchor: 'end',
                        borderColor: 'white',
                        backgroundColor: 'darkgrey',
                        borderRadius: 5,
                        borderWidth: 1,
                        color: 'white',
                        font: { weight: 'bold' },
                        formatter: Math.round,
                        padding: 6
                    }                           
                }
            },
            plugins: [
                {
                    id: 'outerCircleAndLabelsPlugin',
                    afterDatasetsDraw: chart => this.drawOuterCirclePlugin(chart) 
                },
                {
                    id: 'dynamicPadding',
                    afterDraw: (chart) => {
                        const padding = this.calculatePercentagePadding(chart);
                        chart.options.layout.padding = padding;
                    }
                },
                ChartDataLabels
            ],
        });
    }

    private contextMenuStart(){
        // Context menu
        this.canvas.oncontextmenu = (event) => {
            event.preventDefault();
            this.selectionManager = this.host.createSelectionManager();
            const elements = this.chart.getElementsAtEventForMode(event, 'nearest', { intersect: true }, false);
            if (elements.length > 0) {
                const firstElement = elements[0];
                const datasetIndex = firstElement.datasetIndex;
                const dataIndex = firstElement.index;
                const dataPoint = this.chart.data.datasets[datasetIndex].selectionId[dataIndex];
                this.selectionManager.showContextMenu(dataPoint? dataPoint.selection : {}, {
                    x: event.clientX,
                    y: event.clientY
                  });
            }else{
                this.selectionManager.showContextMenu(this.chart.data.datasets[0].selectionId[0]? null: {},{
                    x: event.clientX,
                    y: event.clientY,
                })
            }
        };
    }

    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    //// UPDATE ////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    public update(options: VisualUpdateOptions) {

        this.events.renderingStarted(options);
        
        const dataView: DataView = options.dataViews && options.dataViews[0];
        if (!dataView) return;

        const categories = dataView.categorical.categories;
        if(this.categories_org_order.length === 0){
            this.categories_org_order = categories[0].values;
        }

        // Transform data from Power BI to Chart.js format
        const transformedData = this.transformData(dataView);

        // Categories > for each we'll add the categorySelectionId (after sorting!)
        const originalValues = categories[0].values;
        const sortedLabels = transformedData.labels;
        const originalIdentities = categories[0].identity;
        const sortMapping = new Map();
        sortedLabels.forEach((label, sortedIndex) => {
            const originalIndex = originalValues.findIndex(originalLabel => originalLabel === label);
            sortMapping.set(originalIndex, sortedIndex);
        });
        const sortedIdentities = Array.from(sortMapping.keys()).sort((a, b) => sortMapping.get(a) - sortMapping.get(b)).map(index => originalIdentities[index]);
        categories[0].identity = sortedIdentities;
        categories[0].values = sortedLabels; 

        // const categories = transformedData.order;
        const categoriesCount = categories[0].values.length;
        for (let categoryIndex = 0; categoryIndex < categoriesCount; categoryIndex++) {
            const categoryValue: powerbi.PrimitiveValue = categories[0].values[categoryIndex];
            const categorySelectionId = this.host.createSelectionIdBuilder()
                .withCategory(categories[0], categoryIndex)
                .createSelectionId();
            transformedData.selectionId[categoryIndex] = {
                value: categoryValue,
                selection: categorySelectionId
            };
        }

        transformedData.colorsType = transformedData.colorsType.map(color => color === null ? '#fff' : color);
        transformedData.colors = transformedData.colors.map(color => color === null ? 'rgba(211, 211, 211, 0.5)' : color);

        // Update chart data
        this.chart.data = {
            labels: transformedData.labels,
            datasets: [{
                label: 'Avg.',
                data: transformedData.values.map(value => Math.round(value * 10) / 10), // Round to one decimal place
                secondValues: transformedData.secondValues,
                backgroundColor: transformedData.colors,
                backgroundColorType: transformedData.colorsType,
                dataType: transformedData.types,
                selectionId: transformedData.selectionId
            }]
        };

        // Find the maximum value in the dataset to dynamically set the scale
        const maxDataValue = Math.max(...transformedData.values, ...transformedData.secondValues);
        const maxScaleValue = Math.ceil(maxDataValue / 10) * 10; 

        // Dynamically set the max property of the r scale
        if(maxScaleValue > 100){
            this.chart.options.scales.r.max = maxScaleValue;
            this.chart.options.scales.r.ticks.stepSize = maxScaleValue / 5; 
        }else{
            this.chart.options.scales.r.max = 100;
            this.chart.options.scales.r.ticks.stepSize = 20; 
        }

        this.chart.update();
        setTimeout(() => {  // Cause the Chartjs animation takes 1000ms
            this.events.renderingFinished(options);
        }, 1000);
    }


    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    //// TRANSFORM DATA ////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    private transformData(dataView: DataView) {

        // First update formatting properties
        this.updateFormattingProperties(dataView.metadata.objects, dataView);

        if(dataView.categorical.values[0].values !== this.previous_values){
            if(this.previous_values.length > 0){
                this.categories_org_order = dataView.categorical.values[0].values;  // Data changed, reset categories
            }
            this.previous_values = dataView.categorical.values[0].values;
        }

        let categories;
        try {
            // Retrieve current categories from dataView
            categories = dataView.categorical.categories[0].values.map(value => this.escapeHtml(String(value)));
            // Attempt to sort categories based on the original order in this.categories_org_order
            categories.sort((a, b) => {
                let indexA = this.categories_org_order.indexOf(a);
                let indexB = this.categories_org_order.indexOf(b);

                // Handling categories not found in the original order (-1 comparison result)
                if (indexA === -1) indexA = Number.MAX_SAFE_INTEGER;
                if (indexB === -1) indexB = Number.MAX_SAFE_INTEGER;

                return indexA - indexB;
            });
        } catch (error) {
            // Fallback to original order from dataView if sorting fails
            categories = dataView.categorical.categories[0].values.map(value => this.escapeHtml(String(value)));
        }
        
        const colors = new Array(categories.length).fill(null);
        const colorsType = new Array(categories.length).fill(null);
        let orderValues = new Array(categories.length).fill(null);
        const types = new Array(categories.length).fill(null);
        const firstMeasureValues = new Array(categories.length).fill(null);
        const secondMeasureValues = new Array(categories.length).fill(null); 

        dataView.categorical.values.forEach(valueColumn => {
            if (valueColumn.source.roles.measure) {
                valueColumn.values.forEach((value, index) => {
                    if (firstMeasureValues[index] === null) {
                        firstMeasureValues[index] = value; // Assign to first measure
                    } else {
                        secondMeasureValues[index] = value; // Assign to second measure
                    }
                });
                // Color first measure
                categories.forEach((value, index) => {
                    const category = String(value);
                    categories[index] = category;
                    colors[index] = this.hexToRgba(this.getColorForCategory(category), 0.5);
                });
            } else if (valueColumn.source.roles.order) {
                orderValues = valueColumn.values.map(value => Number(value));
            } else if (valueColumn.source.roles.type) {
                valueColumn.values.forEach((value, index) => {
                    const type = this.escapeHtml(String(value));
                    types[index] = type;
                    colorsType[index] = this.getColorForType(type);
                });
            }
        });

        if (firstMeasureValues.length !== categories.length ||
            secondMeasureValues.length !== categories.length) {
            return; 
        }

        // Combine categories, values, colors, and order into a single array
        const combinedData = categories.map((category, index) => ({
            category: category,
            value: firstMeasureValues[index],
            secondValue: secondMeasureValues[index],
            color: colors[index],
            colorType: colorsType[index],
            type: types[index],
            order: orderValues[index] || 0, // Default to 0 if no order value
            selectionId: {}
        }));

        // Sort the combined data based on order values
        combinedData.sort((a, b) => a.order - b.order);

        // Extract the sorted data back into individual arrays
        let sortedCategories = categories;
        sortedCategories = combinedData.map(item => item.category);            
        const sortedValues = combinedData.map(item => item.value);
        const sortedColors = combinedData.map(item => item.color);
        const sortedColorsType = combinedData.map(item => item.colorType);
        const sortedTypes = combinedData.map(item => item.type);
        const sortedSecondValues = combinedData.map(item => item.secondValue);
        const sortedselectionId = combinedData.map(item => item.selectionId);

        return {
            labels: sortedCategories,
            values: sortedValues,
            selectionId: sortedselectionId,
            colors: sortedColors,
            colorsType: sortedColorsType,
            types: sortedTypes,
            secondValues: sortedSecondValues
        };
    }


    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    //// FORMATTING PANE ////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    private escapeHtml(text) {
        return text.replace(/[&<>"']/g, function(match) {
            switch (match) {
                case '&': return '&amp;';
                case '<': return '&lt;';
                case '>': return '&gt;';
                case '"': return '&quot;';
                case "'": return '&#39;';
                default: return match;
            }
        });
    }

    private updateFormattingProperties(properties: powerbi.DataViewObjects, dataView) {
        
        if (properties && properties.outerCircles) {
            const outerCircleProperties = properties.outerCircles;
            
            this.categoryName = this.escapeHtml(dataView.metadata.columns.find(col => col.roles && col.roles.category).displayName);
            const categories_temp = dataView.categorical.categories[0].values.map(value => this.escapeHtml(String(value)));
            const colorsType = new Array(categories_temp.length).fill(null);
            const types_temp = new Array(categories_temp.length).fill(null);

            // Update color property names based on unique categories
            categories_temp.forEach((category, index) => {
                const propName = `bgColorCategory${index + 1}`; // Dynamically create the property name
                if (Object.prototype.hasOwnProperty.call(outerCircleProperties, propName)) {
                    const bgColorPropertyValue = outerCircleProperties[propName];
                    if (typeof bgColorPropertyValue['solid']['color'] === 'string') {
                        this[propName] = this.escapeHtml(bgColorPropertyValue['solid']['color']);
                    } else {
                        this[propName] = '#d3d3d3'; // Default color
                    }
                    this.colorPaletteInner[index] = this[propName];
                }
            });

            // Check if the font property exists in the formatting properties
            if (Object.prototype.hasOwnProperty.call(outerCircleProperties, 'fontFamily')) {
                const FontTypePropertyValue = this.escapeHtml(outerCircleProperties['fontFamily']);
                if (typeof FontTypePropertyValue === 'string') {
                    this.fontOuterCircle = FontTypePropertyValue;
                } else {
                    this.fontOuterCircle = 'Arial'; // Default font
                }
            }

            // Update color property names based on unique types
            dataView.categorical.values.forEach(valueColumn => {
                if (valueColumn.source.roles.type) {
                    valueColumn.values.forEach((value, index) => {
                        const type = this.escapeHtml(String(value));
                        types_temp[index] = type;
                        colorsType[index] = this.getColorForType(type);
                    });
                }
            });
        
            this.uniqueTypes.forEach((type, index) => {
                const propName = `bgColorType${index + 1}`; // Dynamically create the property name
                if (Object.prototype.hasOwnProperty.call(outerCircleProperties, propName)) {
                    const bgColorPropertyValue = outerCircleProperties[propName];
                    if (typeof bgColorPropertyValue['solid']['color'] === 'string') {
                        this[propName] = this.escapeHtml(bgColorPropertyValue['solid']['color']);
                    } else {
                        this[propName] = '#fff'; // Default color
                    }
                    this.colorPalette[index] = this[propName];
                }
            });
        }
        this.update;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        
        // Building data card, We are going to add two formatting groups "Font Control Group" and "Data Design Group"
        const outerCircles: powerbi.visuals.FormattingCard = {
            // description: "Outer Circle Description",
            displayName: "Outer Circles",
            uid: "outerCircle_uid",
            groups: []
        }

        // Building formatting group "Font Control Group"
        const group1_dataFont: powerbi.visuals.FormattingGroup = {
            displayName: "Font Control Group",
            uid: "outerCircle_fontControl_group_uid",
            slices: [
                {
                    uid: "data_font_control_slice_uid",
                    displayName: "Font",
                    control: {
                        type: powerbi.visuals.FormattingComponent.FontPicker, // Use FontPicker directly
                        properties: {
                            descriptor: {
                                objectName: "outerCircles",
                                propertyName: "fontFamily"
                            },
                            value: "arial, wf_standard-font, helvetica, sans-serif"
                        }
                    }
                }
            ],
        };

        // Building formatting group "Color Group"
        const group2_dataDesign: powerbi.visuals.FormattingGroup = {
            displayName: "Types Colors",
            uid: "outerCircle_dataDesign_group_uid",
            slices: []
        };

        // Generate color slices dynamically based on the unique types
        if(this.uniqueTypes.length !== 0){
            this.uniqueTypes.forEach((type, index) => {
                group2_dataDesign.slices.push({
                    displayName: `${type}`,
                    uid: `outerCircle_dataDesign_bgColor_slice${index + 1}`,
                    control: {
                        type: powerbi.visuals.FormattingComponent.ColorPicker,
                        properties: {
                            descriptor: {
                                objectName: "outerCircles",
                                propertyName: `bgColorType${index + 1}`
                            },
                            value: { value: this[`bgColorType${index + 1}`] }
                        }
                    }
                });
            });
        }

        // Building formatting group "Color Group"
        const group3_dataDesign: powerbi.visuals.FormattingGroup = {
            displayName: "Categories Colors",
            uid: "outerCircle_dataDesign_group2_uid",
            slices: []
        };

        // Generate color slices dynamically based on the unique categories
        if(this.uniqueCategories.length !== 0){
            this.uniqueCategories.forEach((category, index) => { 
                group3_dataDesign.slices.push({
                    displayName: `${category}`,
                    uid: `outerCircle_dataDesign_bgColor2_slice${index + 1}`,
                    control: {
                        type: powerbi.visuals.FormattingComponent.ColorPicker,
                        properties: {
                            descriptor: {
                                objectName: "outerCircles",
                                propertyName: `bgColorCategory${index + 1}`
                            },
                            value: { value: this[`bgColorCategory${index + 1}`] }
                        }
                    }
                });
            });
        }

        // Add formatting groups to data card
        outerCircles.groups.push(group1_dataFont);
        outerCircles.groups.push(group2_dataDesign);
        outerCircles.groups.push(group3_dataDesign);

        // Build and return formatting model with data card
        const formattingModel: powerbi.visuals.FormattingModel = { cards: [outerCircles] };
        return formattingModel;
    }

    private getColorForType(type: string): string {

        // Check if the type is already in the uniqueTypes array
        let index = this.uniqueTypes.indexOf(type);
        
        // If it's a new type, add it to the array
        if (index === -1) {
            this.uniqueTypes.push(type);
            index = this.uniqueTypes.length - 1;
        }
        
        // Return the color corresponding to the type index
        return this.colorPalette[index % this.colorPalette.length];
    }


    private getColorForCategory(category: string): string {

        // Check if the type is already in the uniqueCategories array
        let index = this.uniqueCategories.indexOf(category);
        
        // If it's a new category, add it to the array
        if (index === -1) {
            this.uniqueCategories.push(category);
            index = this.uniqueCategories.length - 1;
        }

        // Return the color corresponding to the type index
        return this.colorPaletteInner[index % this.colorPaletteInner.length];
    }

    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    //// HELPER FUNCTIONS ////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Helper function to convert HEX to RGBA
    private hexToRgba(hex, alpha) {
        // Ensure hex is a string, remove the hash at the start if it's there
        hex = hex.replace(/^#/, '');
        
        // If the code is in shorthand form, convert to full form
        if (hex.length === 3) {
            hex = hex.split('').map(function (hexPart) {
                return hexPart + hexPart;
            }).join('');
        }

        // Convert the r, g, b values to decimal
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);

        // Return the RGBA color string
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }

    // Helper function to convert existing RGBA to new RGBA with a specified alpha
    private adjustRgbaAlpha(rgba, newAlpha) {
        // Use a regular expression to extract the rgba components
        const parts = rgba.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/);

        if (!parts) {
        throw new Error('The provided string does not match the expected RGBA format');
        }

        // Extract the r, g, b values, and use the existing alpha if no new alpha is provided
        const r = parseInt(parts[1], 10);
        const g = parseInt(parts[2], 10);
        const b = parseInt(parts[3], 10);
        const alpha = newAlpha !== undefined ? newAlpha : (parts[4] !== undefined ? parseFloat(parts[4]) : 1);

        // Return the new RGBA color string
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }

    private uniqueTypes: string[] = [];
    private uniqueCategories: string[] = [];

    private colorPalette: string[] = [
        this.bgColorType1,
        this.bgColorType2,
        this.bgColorType3,
        this.bgColorType4,
        this.bgColorType5,
        this.bgColorType6,
        this.bgColorType7,
        this.bgColorType8,
        this.bgColorType9,
        this.bgColorType10
    ];

    private colorPaletteInner: string[] = [
        this.bgColorCategory1,
        this.bgColorCategory2,
        this.bgColorCategory3,
        this.bgColorCategory4,
        this.bgColorCategory5,
        this.bgColorCategory6,
        this.bgColorCategory7,
        this.bgColorCategory8,
        this.bgColorCategory9,
        this.bgColorCategory10,
        this.bgColorCategory11,
        this.bgColorCategory12,
        this.bgColorCategory13,
        this.bgColorCategory14,
        this.bgColorCategory15,
        this.bgColorCategory16,
        this.bgColorCategory17,
        this.bgColorCategory18,
        this.bgColorCategory19,
        this.bgColorCategory20,
        this.bgColorCategory21,
        this.bgColorCategory22,
        this.bgColorCategory23,
        this.bgColorCategory24,
        this.bgColorCategory25,
        this.bgColorCategory26,
        this.bgColorCategory27,
        this.bgColorCategory28,
        this.bgColorCategory29,
        this.bgColorCategory30
    ];
}    
