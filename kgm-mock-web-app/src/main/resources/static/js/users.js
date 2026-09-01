let usersLoaded = false;

function initUsersPanel() {
    if (usersLoaded) return;

    fetch('/api/users')
        .then(response => {
            if (!response.ok) throw new Error("Unauthorized or Error");
            return response.json();
        })
        .then(users => {
            // userId'ye göre küçükten büyüğe sırala
            users.sort((a, b) => a.userId - b.userId);

            const container = document.getElementById('usersList');
            container.innerHTML = '';

            users.forEach(user => {
                const card = document.createElement('div');
                card.className = 'user-card';

                const roleName = user.role ? user.role.name : 'PERSONNEL';
                const isAdmin = roleName === 'ADMIN';
                const initial = user.userFullname ? user.userFullname.charAt(0).toUpperCase() : 'U';
                const email = user.userEmail || 'No email provided';

                card.innerHTML = `
                    <div class="user-avatar-wrapper">
                        <img class="user-avatar-img" 
                             src="/api/users/${user.userId}/avatar" 
                             alt="${user.userFullname}" 
                             onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';">
                        <div class="user-avatar-placeholder" style="display: none;">${initial}</div>
                    </div>
                    <span class="user-role-badge ${isAdmin ? 'admin' : ''}">${roleName}</span>
                    <div class="user-fullname" title="${user.userFullname}">${user.userFullname}</div>
                    <div class="user-email">${email}</div>
                `;

                container.appendChild(card);
            });

            usersLoaded = true;
        })
        .catch(err => {
            console.error("Failed to load users:", err);
        });
}