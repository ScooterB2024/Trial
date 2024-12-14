document.addEventListener('DOMContentLoaded', () => {
    const reportForm = document.getElementById('reportForm');
    const findingsContainer = document.getElementById('findingsContainer');
    const addFindingButton = document.getElementById('addFinding');
    const reportOutput = document.getElementById('reportOutput');
    const exportButton = document.getElementById('exportButton');

    let alarmChart;

    addFindingButton.addEventListener('click', () => {
        const newFinding = document.createElement('div');
        newFinding.classList.add('finding');
        newFinding.innerHTML = `
            <div class="finding-top">
                <label for="machineId">Machine ID:</label>
                <input type="text" class="machineId" required>
                <label for="description">Description:</label>
                <textarea class="description" required></textarea>
            </div>
            <label for="recommendedAction">Recommended Action:</label>
            <textarea class="recommendedAction" required></textarea>
            <label for="image">Image:</label>
            <input type="file" class="image" accept="image/*">
        `;
        findingsContainer.appendChild(newFinding);
    });

    reportForm.addEventListener('submit', (e) => {
        e.preventDefault();
        generateReport();
    });

    function generateReport() {
        const machinesMonitored = document.getElementById('machinesMonitored').value;
        const highAlarm = document.getElementById('highAlarm').value;
        const mediumAlarm = document.getElementById('mediumAlarm').value;
        const lowAlarm = document.getElementById('lowAlarm').value;
        const notAvailable = document.getElementById('notAvailable').value;
        const analyst = document.getElementById('analyst').value;
        const dateOfDataCollection = document.getElementById('dateOfDataCollection').value;
        const dateOfReview = document.getElementById('dateOfReview').value;

        let findings = [];
        const findingElements = document.querySelectorAll('.finding');
        findingElements.forEach(findingElement => {
            const machineId = findingElement.querySelector('.machineId').value;
            const description = findingElement.querySelector('.description').value;
            const recommendedAction = findingElement.querySelector('.recommendedAction').value;
            const imageFile = findingElement.querySelector('.image').files[0];
            let imageUrl = '';
            if (imageFile) {
                imageUrl = URL.createObjectURL(imageFile);
            }
            findings.push({ machineId, description, recommendedAction, imageUrl });
        });

        let reportContent = `
            <h2>Vibration Analysis Report</h2>
            <p>Date of Data Collection: ${dateOfDataCollection}</p>
            <p>Date of Review: ${dateOfReview}</p>
            <p>Analyst: ${analyst}</p>
            <h3>Summary</h3>
            <p>Machines Monitored: ${machinesMonitored}</p>
            <p>High Alarm: ${highAlarm}</p>
            <p>Medium Alarm: ${mediumAlarm}</p>
            <p>Low Alarm: ${lowAlarm}</p>
            <p>Not Available: ${notAvailable}</p>
            <h3>Findings</h3>
        `;

        findings.forEach(finding => {
            reportContent += `
                <div class="finding-report">
                    <p>Machine ID: ${finding.machineId}</p>
                    <p>Description: ${finding.description}</p>
                    <p>Recommended Action: ${finding.recommendedAction}</p>
                    ${finding.imageUrl ? `<img src="${finding.imageUrl}" alt="Finding Image" style="max-width: 200px; max-height: 200px;">` : ''}
                </div>
            `;
        });

        reportOutput.innerHTML = reportContent;
        reportOutput.style.display = 'block';
        exportButton.style.display = 'block';

        generateAlarmChart(machinesMonitored, highAlarm, mediumAlarm, lowAlarm, notAvailable);
    }

    function generateAlarmChart(machinesMonitored, highAlarm, mediumAlarm, lowAlarm, notAvailable) {
        const ctx = document.getElementById('alarmChart').getContext('2d');
        if (alarmChart) {
            alarmChart.destroy();
        }
        alarmChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['High Alarm', 'Medium Alarm', 'Low Alarm', 'Not Available'],
                datasets: [{
                    label: 'Number of Alarms',
                    data: [highAlarm, mediumAlarm, lowAlarm, notAvailable],
                    backgroundColor: ['#ff0000', '#ffa500', '#ffff00', '#d3d3d3']
                }]
            },
            options: {
                responsive: true,
                scales: {
                    y: {
                        beginAtZero: true
                    }
                }
            }
        });
    }

    exportButton.addEventListener('click', () => {
        exportToWord();
    });

    function exportToWord() {
        const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>";
        const bodyStart = "<body>";
        const bodyEnd = "</body></html>";
        const htmlContent = header + bodyStart + reportOutput.innerHTML + bodyEnd;

        const blob = new Blob(['\ufeff', htmlContent], {
            type: 'application/msword'
        });

        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'Vibration_Analysis_Report.doc';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
});
