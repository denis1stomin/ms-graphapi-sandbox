const GRAPHAPI_TOKEN = process.env.GRAPHAPI_TOKEN;

const Argv = require('minimist')(process.argv.slice(2));
const MicrosoftGraph = require("@microsoft/microsoft-graph-client");

const Client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
        // first parameter takes an error if you can't get an access token
        done(null, GRAPHAPI_TOKEN);
    }
});

function getFilePath() {
    return '/me/onenote/pages/0-fccfdd92d7074df1a2dab913541b9df4!8-3F1121B161CB5087!121/content';
}

function saveText(text) {
    Client
        .api(getFilePath())
        .header('Content-Type', 'text/html')
        .patch(text, (err, res) => {
            if (err) {
                console.log(err);
                process.exit(1);
            }
        });
}

function readText() {
    Client
        .api(getFilePath())
        .get((err, res, raw) => {
            if (raw) {
                console.log(raw.text);
            }
            else {
                console.log(err);
                process.exit(1);
            }
        });
}

Client
    .api('/me')
    .get((err, res) => {
        if (res) {
            console.log(`Working on behalf of ${res.displayName}`);
        }
        else {
            console.log(err);
            process.exit(1);
        }
    });

var text2save = Argv['text'];
if (text2save) {
    saveText(text2save);
}
else {
    readText();
}
