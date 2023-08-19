if (window.location.host != "teams.microsoft.com") {
    window.location.replace("https://teams.microsoft.com/")
}

fetch('https://cdn.jsdelivr.net/gh/brix/crypto-js/crypto-js.min.js')
    .then(response => eval(response.text()))

var messagesElement = document.createElement('div');
messagesElement.id = "chatMessages";
messagesElement.style.overflow = 'auto';
messagesElement.className = window.frames[0].document.querySelector('[data-tid="message-pane-list-runway"]').className
window.frames[0].document.querySelector('[data-tid="message-pane-body"]').replaceChild(messagesElement, window.frames[0].document.querySelector('[data-acc-id="announcing-region-message-list"]'));

var inputElement = document.createElement('textArea');
inputElement.id = "inputBox";
inputElement.style.width = '100%';
inputElement.style.height = '100%';
inputElement.style.backgroundColor = '#292929';
inputElement.onkeyup = function(event) {
    if (event.keyCode == 13 && !event.shiftKey) {
        sendMessage(`[encryTeams]${encrypt(window.frames[0].document.getElementById('inputBox').value, encryptionKeys[currentChatId])}`)
        window.frames[0].document.getElementById('inputBox').value = '';
    }
};
window.frames[0].document.querySelector('[data-tid="message-pane-footer"]').replaceChild(inputElement, window.frames[0].document.querySelector('[data-tid="chat-pane-compose-message-footer"]'));

fetch('https://cdn.jsdelivr.net/gh/brix/crypto-js/crypto-js.min.js')
    .then(response => eval(response.text()))

function createSetKeyButton() {
    if (!document.getElementById('setKey')) {
        var setKey = document.createElement('button');
        setKey.id = "setKey";
        setKey.innerText = "Set Key";
        setKey.style.backgroundColor = '#1F1F1F';
        setKey.style.color = '#FFFFFF';
        document.querySelector('[data-tid="tabs-menu"]').append(setKey);
        document.getElementById('setKey').onclick = function() {
            encryptionKeys[currentChatId] = prompt("Enter Key:");
            localStorage.setItem("AESkeys", JSON.stringify(encryptionKeys));
        }
    }
}
createSetKeyButton()

var username = document.getElementsByClassName("user-picture")[0].getAttribute('alt').split('f ')[1].replace('.', '');

var teamsToken;
Object.keys(localStorage).forEach(e => {
    try {
        var r = JSON.parse(localStorage.getItem(e)).skypeToken;
        if (r != undefined) {
            teamsToken = r;
        }
    } catch {}
});

var currentChatId = window.location.hash.split('/')[2].split('?')[0]
window.onhashchange = function() {
    currentChatId = window.location.hash.split('/')[2].split('?')[0]
    getMessages();
    createSetKeyButton()
};

var encryptionKeys;
if (!localStorage.getItem("AESkeys")) {
    localStorage.setItem("AESkeys", JSON.stringify({}));
}
encryptionKeys = JSON.parse(localStorage.getItem("AESkeys"));

function encrypt(plain, password) {
    return CryptoJS.AES.encrypt(plain, password).toString()
}

function decrypt(encrypted, password) {
    try {
        var decrypted = CryptoJS.AES.decrypt(encrypted, password).toString(CryptoJS.enc.Utf8);
    } catch {
        var decrypted = false;
    }
    if (encrypted && !decrypted) {
        return "INVALID KEY";
    } else {
        return decrypted;
    }
}

function sendMessage(message) {
    let xhr = new XMLHttpRequest();
    xhr.open('POST', `https://amer.ng.msg.teams.microsoft.com/v1/users/ME/conversations/${currentChatId}/messages`);
    xhr.setRequestHeader('authentication', 'skypetoken=' + teamsToken)
    xhr.send(JSON.stringify({
        "content": message,
        "messagetype": "text",
        "imdisplayname": username
    }))
}

var previousMessages = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];
var failCount = 0;

function getMessages() {
    let xhr = new XMLHttpRequest();
    xhr.open('GET', `https://amer.ng.msg.teams.microsoft.com/v1/users/ME/conversations/${currentChatId}/messages?startTime=0&pageSize=50`);
    xhr.setRequestHeader('authentication', 'skypetoken=' + teamsToken)
    xhr.onload = function() {
        if (xhr.status == 200) {
            failCount = 0;
            var rawMessages = JSON.parse(xhr.response).messages;
            var isSame = true;
            rawMessages.forEach((message, index) => {
                if (message.content != previousMessages[index]) {
                    isSame = false;
                }
            });
            if (isSame) {
                return;
            } else {
                clearTimeout(stopUpdating);
                stopUpdating = setTimeout(window.location.reload, 1000 * 60 * 5);
                rawMessages.forEach((message, index) => {
                    previousMessages[index] = message.content;
                });
                displayMessages(rawMessages);
            }
        } else {
            if (failCount) {
                window.location.reload();
            }
            failCount++;
            Object.keys(localStorage).forEach(e => {
                try {
                    var r = JSON.parse(localStorage.getItem(e)).skypeToken;
                    if (r != undefined) {
                        teamsToken = r;
                    }
                } catch {}
            });
            getMessages();
        }
    }
    xhr.send()
}

function parseMessage(message) {
    if (message.substring(0, 12) == "[encryTeams]") {
        return decrypt(message.substring(12), encryptionKeys[currentChatId]).replace(/\n/g, '<br>');
    } else {
        return message;
    }
}

function displayMessages(messages) {
    window.frames[0].document.getElementById('chatMessages').innerHTML = '';
    var lastMessageAuthor;

    for (var i = 0; i < messages.length; i++) {
        var messageAuthor = messages[messages.length - i - 1].imdisplayname;
        var messageContent = parseMessage(messages[messages.length - i - 1].content);
        if (!messageContent) {
            continue;
        }

        var messageElement = document.createElement('div');
        messageElement.style.color = '#989898';
        messageElement.style.fontSize = '12px';
        var messageTime = new Date(messages[messages.length - i - 1].originalarrivaltime)

        var messageContentElement = document.createElement('div');
        messageContentElement.style.color = '#FFFFFF';
        messageContentElement.style.fontSize = '14px';
        messageContentElement.style.padding = '6px 15px 8px 15px';
        messageContentElement.style.borderRadius = '6px';
        messageContentElement.style.wordBreak = 'break-word';
        messageContentElement.style.width = 'fit-content';
        messageContentElement.innerHTML = messageContent;

        if (messageAuthor == username) {
            messageElement.innerHTML = `${messageTime.toLocaleDateString()} ${messageTime.getHours()}:${messageTime.getMinutes().length == 1 ? '0' + messageTime.getMinutes() : messageTime.getMinutes()}`;
            messageElement.style.textAlign = 'right';
            messageElement.style.marginLeft = 'auto';
            messageElement.style.marginRight = '0px';
            messageContentElement.style.backgroundColor = '#2B2B40';
            messageElement.appendChild(messageContentElement);
            window.frames[0].document.getElementById('chatMessages').appendChild(messageElement);
        } else {
            messageContentElement.style.backgroundColor = '#292929';
            if (messageAuthor == lastMessageAuthor) {
                messageElement.innerHTML = `${messageTime.toLocaleDateString()} ${messageTime.getHours()}:${messageTime.getMinutes()}`;
                messageElement.appendChild(messageContentElement);
                window.frames[0].document.getElementById('chatMessages').appendChild(messageElement);
            } else {
                var PFPcontainer = document.createElement('div');
                PFPcontainer.style.display = 'flex';
                PFPcontainer.style.position = 'relative';
                PFPcontainer.style.left = '-32px';
                var img = document.createElement('img');
                img.src = `https://teams.microsoft.com/api/mt/amer/beta/users/${messages[messages.length - i - 1].from.split('/')[7]}/profilepicturev2?displayname=${encodeURIComponent(messages[messages.length - i - 1].imdisplayname)}&size=HR64x64`
                img.style.width = '32px';
                img.style.height = '32px';
                img.style.borderRadius = '50%';
                img.style.position = 'relative';
                img.style.top = '20px';
                var messageContentContainer = document.createElement('div');
                messageContentContainer.innerHTML = `${messages[messages.length - i - 1].imdisplayname} ${messageTime.toLocaleDateString()} ${messageTime.getHours()}:${messageTime.getMinutes()}`;
                messageContentContainer.appendChild(messageContentElement);
                PFPcontainer.appendChild(img);
                PFPcontainer.appendChild(messageContentContainer);
                window.frames[0].document.getElementById('chatMessages').appendChild(PFPcontainer);
            }
        }
        lastMessageAuthor = messageAuthor;
    }
    var chatWindow = window.frames[0].document.getElementById('chatMessages');
    chatWindow.scrollTop = chatWindow.scrollHeight;
}

setInterval(getMessages, 3000);
var stopUpdating = setTimeout(window.location.reload, 1000 * 60 * 5);
