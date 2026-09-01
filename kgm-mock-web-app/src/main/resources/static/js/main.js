document.addEventListener('DOMContentLoaded', () => {
    startClockAndShiftTracker();
    loadDashboardStats();
});

function startClockAndShiftTracker() {
    function update() {
        const now = new Date();

        // Saat (HH:MM:SS)
        const hrs = String(now.getHours()).padStart(2, '0');
        const mins = String(now.getMinutes()).padStart(2, '0');
        const secs = String(now.getSeconds()).padStart(2, '0');
        document.getElementById('liveClock').textContent = `${hrs}:${mins}:${secs}`;

        // Tarih (Örn: 24 MAY 2026, SUNDAY)
        const dateOptions = { day: 'numeric', month: 'short', year: 'numeric', weekday: 'long' };
        document.getElementById('liveDate').textContent = now.toLocaleDateString('en-GB', dateOptions).toUpperCase();

        // Vardiya Hesaplama: 3 Vardiya (00:00-08:00, 08:00-16:00, 16:00-00:00)
        const currentHour = now.getHours();
        const currentMin = now.getMinutes();
        const currentSec = now.getSeconds();

        let shiftNum = 1;
        let shiftName = "Night Shift (00:00 - 08:00)";
        let startHour = 0;
        let endHour = 8;

        if (currentHour >= 8 && currentHour < 16) {
            shiftNum = 2;
            shiftName = "Day Shift (08:00 - 16:00)";
            startHour = 8;
            endHour = 16;
        } else if (currentHour >= 16 && currentHour < 24) {
            shiftNum = 3;
            shiftName = "Evening Shift (16:00 - 00:00)";
            startHour = 16;
            endHour = 24;
        }

        // Aktif vardiya slotunu vurgula
        document.querySelectorAll('.timeline-slot').forEach(s => s.classList.remove('active'));
        const activeSlot = document.getElementById(`slot-${shiftNum}`);
        if (activeSlot) activeSlot.classList.add('active');

        document.getElementById('currentShiftName').textContent = shiftName;

        // Vardiyada geçen süre & Kalan süre hesabı
        const totalShiftSeconds = 8 * 3600;
        const elapsedSeconds = ((currentHour - startHour) * 3600) + (currentMin * 60) + currentSec;
        const remainingSeconds = totalShiftSeconds - elapsedSeconds;

        const progressPercent = Math.min(100, Math.max(0, (elapsedSeconds / totalShiftSeconds) * 100));
        document.getElementById('shiftProgressBar').style.width = `${progressPercent.toFixed(1)}%`;

        const remHours = Math.floor(remainingSeconds / 3600);
        const remMins = Math.floor((remainingSeconds % 3600) / 60);
        document.getElementById('shiftRemainingTime').textContent = `Next shift in: ${remHours}h ${remMins}m`;
    }

    update();
    setInterval(update, 1000);
}

function loadDashboardStats() {
    fetch('/api/users')
        .then(res => res.json())
        .then(users => {
            const countElem = document.getElementById('totalUsersCount');
            if (countElem) {
                countElem.textContent = users.length;
            }
        })
        .catch(() => {
            const countElem = document.getElementById('totalUsersCount');
            if (countElem) countElem.textContent = '5'; // Hata durumunda mock değer
        });
}