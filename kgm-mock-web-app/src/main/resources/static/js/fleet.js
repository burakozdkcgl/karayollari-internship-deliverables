let fleetChartInstance = null;
let cachedStatuses = [];
let cachedRegions = [];

function initFleetPanel() {
    Promise.all([
        fetch('/api/fleet/statuses').then(r => r.json()),
        fetch('/api/fleet/regions').then(r => r.json())
    ]).then(([statuses, regions]) => {
        cachedStatuses = statuses;
        cachedRegions = regions;
        loadFleetData();
        loadFleetStats();
    }).catch(err => console.error(err));
}

function loadFleetData() {
    fetch('/api/fleet')
        .then(response => response.json())
        .then(data => {
            const tbody = document.getElementById('fleetTableBody');
            tbody.innerHTML = '';
            
            data.forEach(vehicle => {
                const tr = document.createElement('tr');

                // Status Options
                let statusOptions = cachedStatuses.map(s => 
                    `<option value="${s.id}" ${s.id === vehicle.status.id ? 'selected' : ''}>${s.statusCode}</option>`
                ).join('');

                // Region Options
                let regionOptions = cachedRegions.map(r => 
                    `<option value="${r.id}" ${r.id === vehicle.region.id ? 'selected' : ''}>${r.name}</option>`
                ).join('');

                tr.innerHTML = `
                    <td><strong>${vehicle.plateNumber}</strong></td>
                    <td>${vehicle.type}</td>
                    <td>
                        <select class="table-select" onchange="updateVehicle(${vehicle.id}, this.value, document.getElementById('reg-${vehicle.id}').value)">
                            ${statusOptions}
                        </select>
                    </td>
                    <td>
                        <select id="reg-${vehicle.id}" class="table-select" onchange="updateVehicle(${vehicle.id}, this.previousElementSibling ? this.previousElementSibling.value : this.parentElement.previousElementSibling.querySelector('select').value, this.value)">
                            ${regionOptions}
                        </select>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        });
}

function updateVehicle(vehicleId, statusId, regionId) {
    fetch(`/api/fleet/${vehicleId}?statusId=${statusId}&regionId=${regionId}`, {
        method: 'PUT'
    }).then(res => {
        if (res.ok) {
            loadFleetStats();
        }
    });
}

function loadFleetStats() {
    fetch('/api/fleet/stats')
        .then(response => response.json())
        .then(stats => {
            const labels = Object.keys(stats);
            const values = Object.values(stats);
            const ctx = document.getElementById('fleetChart').getContext('2d');

            if (fleetChartInstance) {
                fleetChartInstance.destroy();
            }

            fleetChartInstance = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: labels,
                    datasets: [{
                        data: values,
                        backgroundColor: ['#1a1816', '#a3382b', '#8c8275']
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { position: 'bottom' }
                    }
                }
            });
        });
}