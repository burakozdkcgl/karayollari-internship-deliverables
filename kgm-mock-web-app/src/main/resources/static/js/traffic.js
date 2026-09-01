const TURKEY_CITIES = [
    { name: "Adana", lat: 37.0000, lng: 35.3213 },
    { name: "Ankara", lat: 39.9334, lng: 32.8597 },
    { name: "Antalya", lat: 36.8969, lng: 30.7133 },
    { name: "Bolu", lat: 40.7350, lng: 31.6061 },
    { name: "Bursa", lat: 40.1885, lng: 29.0610 },
    { name: "Diyarbakir", lat: 37.9144, lng: 40.2306 },
    { name: "Erzurum", lat: 39.9055, lng: 41.2658 },
    { name: "Gaziantep", lat: 37.0662, lng: 37.3833 },
    { name: "Istanbul", lat: 41.0082, lng: 28.9784 },
    { name: "Izmir", lat: 38.4237, lng: 27.1428 },
    { name: "Kayseri", lat: 38.7205, lng: 35.4826 },
    { name: "Konya", lat: 37.8746, lng: 32.4932 },
    { name: "Samsun", lat: 41.2867, lng: 36.3300 },
    { name: "Trabzon", lat: 41.0027, lng: 39.7168 }
];

let map = null;
let currentRouteLayer = null;
let markers = [];
let isMapInitialized = false;

function initRoadMap() {
    const originSelect = document.getElementById('originCity');
    const destSelect = document.getElementById('destCity');

    if (!isMapInitialized) {
        // Dropdown listelerini doldur
        TURKEY_CITIES.forEach((city, index) => {
            const opt1 = new Option(city.name, index);
            const opt2 = new Option(city.name, index);
            originSelect.add(opt1);
            destSelect.add(opt2);
        });

        // Haritayı başlat
        map = L.map('kgmMap').setView([39.0, 35.0], 6);

        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            maxZoom: 18,
            attribution: '© OpenStreetMap'
        }).addTo(map);

        isMapInitialized = true;
    }

    // Tab geçişinde Leaflet boyutunu düzelt
    setTimeout(() => {
        if (map) map.invalidateSize();
    }, 200);
}

function calculateRoute() {
    const originIdx = document.getElementById('originCity').value;
    const destIdx = document.getElementById('destCity').value;

    if (originIdx === "" || destIdx === "") {
        alert("Please select both origin and destination cities.");
        return;
    }

    if (originIdx === destIdx) {
        alert("Origin and destination cannot be the same.");
        return;
    }

    const c1 = TURKEY_CITIES[originIdx];
    const c2 = TURKEY_CITIES[destIdx];

    if (currentRouteLayer) map.removeLayer(currentRouteLayer);
    markers.forEach(m => map.removeLayer(m));
    markers = [];

    const url = `https://router.project-osrm.org/route/v1/driving/${c1.lng},${c1.lat};${c2.lng},${c2.lat}?overview=full&geometries=geojson`;

    fetch(url)
        .then(res => res.json())
        .then(data => {
            if (!data.routes || data.routes.length === 0) return;

            const route = data.routes[0];
            const distanceKm = (route.distance / 1000).toFixed(1);
            
            const totalMin = Math.round(route.duration / 60);
            const hours = Math.floor(totalMin / 60);
            const mins = totalMin % 60;
            const durationStr = `${hours}h ${mins}m`;

            const routeGeoJSON = route.geometry;
            currentRouteLayer = L.geoJSON(routeGeoJSON, {
                style: { color: '#a3382b', weight: 4, opacity: 0.85 }
            }).addTo(map);

            const m1 = L.marker([c1.lat, c1.lng]).addTo(map).bindPopup(`Origin: ${c1.name}`);
            const m2 = L.marker([c2.lat, c2.lng]).addTo(map).bindPopup(`Destination: ${c2.name}`);
            markers.push(m1, m2);

            map.fitBounds(currentRouteLayer.getBounds(), { padding: [40, 40] });

            document.getElementById('resDistance').textContent = `${distanceKm} km`;
            document.getElementById('resDuration').textContent = durationStr;
            document.getElementById('routeResults').style.display = 'flex';
        })
        .catch(err => {
            console.error("Routing error:", err);
            alert("Route calculation failed.");
        });
}