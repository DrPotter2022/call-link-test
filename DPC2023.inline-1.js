const dataMap = {
    "胃の悪性新生物": [0, 0, null, 0, null, 0, 0, 0, 0, 0.1, 0.2, 0.2, 0.5, 0.8, 1.4, 1.7, 1.7, 1.1, 0.7],
    "結腸の悪性新生物": [null, null, null, 0, 0, 0, 0, 0, 0.1, 0.1, 0.3, 0.4, 0.5, 0.9, 1.6, 1.9, 1.8, 1.4, 1.0],
    "肺の悪性新生物": [null, null, null, null, null, null, 0, 0, 0, 0.1, 0.4, 0.5, 1.0, 1.6, 3.1, 3.2, 2.2, 1.3, 0.8],
    "乳房の悪性新生物": [null, null, 0, null, null, null, 0, 0.1, 0.1, 0.5, 0.5, 0.5, 0.5, 0.6, 0.7, 0.6, 0.4, 0.3, 0.3],
    "前立腺の悪性新生物": [0, null, null, null, 0, null, null, null, null, 0, 0, 0.1, 0.2, 0.6, 1.0, 1.0, 0.8, 0.5, 0.4]
};

const labels = [
    "０歳", "１～４歳", "５～９歳", "１０～１４歳", "１５～１９歳", "２０～２４歳",
    "２５～２９歳", "３０～３４歳", "３５～３９歳", "４０～４４歳", "４５～４９歳",
    "５０～５４歳", "５５～５９歳", "６０～６４歳", "６５～６９歳", "７０～７４歳",
    "７５～７９歳", "８０～８４歳", "８５～８９歳", "９０歳以上"
];

const ctx = document.getElementById('riskChart').getContext('2d');
let chart;

function updateChart(disease) {
    if (chart) chart.destroy();

    if (!disease || !dataMap[disease]) {
        document.getElementById('chart-container').style.display = 'none';
        return;
    }

    document.getElementById('chart-container').style.display = 'block';

    chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: `${disease}の罹患率（％）`,
                data: dataMap[disease],
                fill: false,
                borderColor: 'blue',
                tension: 0.1
            }]
        },
        options: {
            responsive: true,
            plugins: {
                tooltip: {
                    enabled: true,
                    callbacks: {
                        label: function(context) {
                            return `${context.parsed.y}%`;
                        }
                    }
                },
                legend: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: '罹患率（％）'
                    }
                }
            }
        }
    });
}

document.getElementById('disease-select').addEventListener('change', function() {
    updateChart(this.value);
});
