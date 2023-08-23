const msalConfig = {
    auth: {
        clientId: '9d3b24e1-b05c-480d-b85b-42d1c7f7a01a', // Replace with your actual client ID
        redirectUri:'http://localhost:8080'
    },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const tokenRequest = {
    scopes: ['user.read', 'mail.read'],
};

msalInstance.loginPopup(tokenRequest)
    .then(response => {
        const accessToken = response.accessToken;
        getEmails(accessToken);
    })
    .catch(error => {
        console.error(error);
    });

function getEmails(accessToken) {
    const emailListElement = document.getElementById('emailList');

    fetch('https://graph.microsoft.com/v1.0/me/messages', {
        headers: {
            Authorization: `Bearer ${accessToken}`,
        },
    })
        .then(response => response.json())
        .then(data => {
            const emails = data.value;
            emails.forEach(email => {
                const emailItem = document.createElement('div');
                emailItem.className = 'email';
                emailItem.innerHTML = `<strong>${email.sender.emailAddress.name}</strong>: ${email.subject}`;
                emailListElement.appendChild(emailItem);
            });
        })
        .catch(error => {
            console.error(error);
        });
}