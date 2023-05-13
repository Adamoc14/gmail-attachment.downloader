import { tokenClient } from "./gapi.js";

/**
 *  Sign in the user upon button click.
 */
const handleAuthClick = (cb) => {
    tokenClient.callback = async (resp) => {
        if (resp.error !== undefined) {
            throw (resp);
        }
        document.getElementById('signout_button').style.visibility = 'visible';
        document.getElementById('authorize_button').innerText = 'Refresh';
        displayErr({ message: '' })
        cb()
    };

    if(gapi.client.getToken() === null) {
        // Prompt the user to select a Google Account and ask for consent to share their data
        // when establishing a new session.
        tokenClient.requestAccessToken({prompt: 'consent'});
    } else {
        // Skip display of account chooser and consent dialog for an existing session.
        tokenClient.requestAccessToken({prompt: ''});
    }
}

    /**
     *  Sign out the user upon button click.
     */
const handleSignoutClick = ({content_area , authorizeBtn, authenticateBtn }) => {
    const token = gapi.client?.getToken();
    if (token !== null) {
        google.accounts.oauth2.revoke(token.access_token);
        gapi.client.setToken('');
        resetAuthBtns(content_area , authorizeBtn, authenticateBtn)
    }
}


const resetAuthBtns = (content_area , authorizeBtn, authenticateBtn) => {
    content_area.innerText = '';
    authorizeBtn.innerText = 'Authorize';
    authenticateBtn.style.visibility = 'hidden';
}

const displayErr = ({message}) => {
    document.getElementById('content').innerText = message;
}

const decodeBase64 = ({inputStr}) => {
    var binaryString = window.atob(inputStr);
    var bytes = new Uint8Array(binaryString.length);
    for (var i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
}

const obtainMsgResults = async({uId, q}) => {
    let err= new Error();
    const {result: queryResults } = await gapi.client.gmail.users.messages.list({
        'userId': uId,
        "q": q
    });
    const { messages } = queryResults
    if (!messages || messages?.length == 0) err.message = 'Invalid request: no results found, please enter a valid voucher number'
    else if (messages.length > 1) err.message = 'You have entered a voucher number that has too many email results';
    return err.message ? err : messages
}

const obtainMsgSingle = async({uId, msgResults, idx}) => {
    const {id: messageId} = msgResults[idx]
    const {result: msgQuery} = await gapi.client.gmail.users.messages.get({
        'userId': uId,
        "id": messageId
    });
    return msgQuery
}

const downloadAsFile = ({byteArrayBuffer, filename}) => {
    const mimeType = 'application/pdf'
    let blob = new Blob([byteArrayBuffer], {type: mimeType});

    // Back Tracking just to support IE and other outdated browsers (Backwards Compatability)
    if (window.navigator.msSaveOrOpenBlob)  window.navigator.msSaveBlob(blob, filename);
    else {
        let a = window.document.createElement("a");
        a.href = window.URL.createObjectURL(blob, {type: mimeType});
        a.download = filename;
        document.body.appendChild(a);
        a.click();  
        document.body.removeChild(a);
    } 
}

const sanitizeData = ({str}) => {
    const replacements = {
        ' ': '+',
        '_': '/',
        '-': '+'
    }
    for (const [key, value] of Object.entries(replacements)){
        str = str.replace(new RegExp(key, 'g'), value)
    }
    return str
}


const getAttachment = async({uId, msg, idx}) => {
    const err = new Error()
    const {id , payload:data} = msg
    const attachments = data.parts?.[idx] 
    if(!attachments) err.message ='No attachments could be found for this email'
    
    const attachmentID =  attachments.body?.attachmentId
    const filename = attachments.filename
    const {result: attachmentsQuery} = await gapi.client.gmail.users.messages.attachments.get({
        'userId': uId,
        "messageId": id,
        "id": attachmentID
    }); 
    return err.message ? err : {
        filename,
        dataStr: attachmentsQuery.data
    }
}

const downloadVoucher = async({ voucherNumber }) => {
    const userId = 'me'
    const query = voucherNumber
    const ogFilename = `${query}.pdf`
    try {

      // Obtain email results for query (ie. voucher number) passed in   
      const msgResponse = await obtainMsgResults({uId: userId, q: query})
      if (msgResponse instanceof Error) return displayErr({message: msgResponse.message})
      else {

            // Obtain single email result based on parameters passed
            const msgQuery = await obtainMsgSingle({uId: userId, msgResults: msgResponse, idx: 0 }) 
            if (!msgQuery) return displayErr({message: 'Invalid request'})

            // Obtain attachment from message provided
            const attResponse = await getAttachment({uId: userId, msg: msgQuery, idx: 1, })
            if(attResponse instanceof Error) return displayErr({message: attResponse.message})
            else {
                const {filename , dataStr} = attResponse
                const updatedFilename = filename ??  ogFilename
                
                // Santitize output string and download from ArrayBufffer Data
                const sanitizedStr = sanitizeData({str: dataStr})
                const decodedArrayBuffer = decodeBase64({inputStr: sanitizedStr})
                downloadAsFile({byteArrayBuffer: decodedArrayBuffer, filename: ogFilename})
            }
        };
    } catch (err) {
       return displayErr({ message: err.message });
    }
}

const setUpListeners = () => {
    const downloadVoucherBtn = document.querySelector('[role="button"][name="download-voucher-btn"]');
    downloadVoucherBtn.onclick = () => {
        const voucherNumber = document.querySelector('[name="gmail-search-input"]')?.value
        if(voucherNumber.length === 0) return displayErr({message: 'No input was provided'})
        else {
            downloadVoucher({voucherNumber});
        }
    }
}

const assignFns = () => {
    const [content_area, authorizeBtn, authenticateBtn] = [
        document.querySelector('#content'),
        document.querySelector('#authorize_button'),
        document.querySelector('#signout_button')
    ]
    document.querySelector('#authorize_button').onclick = () => handleAuthClick(setUpListeners);
    document.querySelector('#signout_button').onclick = handleSignoutClick(content_area, authorizeBtn, authenticateBtn);
}

export {
    assignFns,
    displayErr
}
