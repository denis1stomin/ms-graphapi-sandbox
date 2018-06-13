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
    return '/me/drive/root:/Documents/mydata.txt:/content';
}

function saveText(text) {
    let ostream = new require('stream').Readable();
    ostream._read = function noop() {}; // some work-around here
    ostream.push(text);
    ostream.push(undefined);

    Client
        .api(getFilePath())
        .put(ostream, (err) => {
            console.log(err);
        });
}

function readText() {
    Client
        .api(getFilePath())
        .getStream((err, downloadStream) => {
            downloadStream.pipe(process.stdout);
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
