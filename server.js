require('dotenv').config() 
var express = require('express')
var request = require('request')
var app = express()

var accessToken = ""

const port = 3000

app.get('/', (req, res) => res.send('Hello World!'))

app.get('/auth/redirect', (req, res) =>{
//     POST /{tenant}/oauth2/token HTTP/1.1
// Host: https://login.microsoftonline.com
// Content-Type: application/x-www-form-urlencoded
// grant_type=authorization_code
// &client_id=6731de76-14a6-49ae-97bc-6eba6914391e
// &code=AwABAAAAvPM1KaPlrEqdFSBzjqfTGBCmLdgfSTLEMPGYuNHSUYBrqqf_ZT_p5uEAEJJ_nZ3UmphWygRNy2C3jJ239gV_DBnZ2syeg95Ki-374WHUP-i3yIhv5i-7KU2CEoPXwURQp6IVYMw-DjAOzn7C3JCu5wpngXmbZKtJdWmiBzHpcO2aICJPu1KvJrDLDP20chJBXzVYJtkfjviLNNW7l7Y3ydcHDsBRKZc3GuMQanmcghXPyoDg41g8XbwPudVh7uCmUponBQpIhbuffFP_tbV8SNzsPoFz9CLpBCZagJVXeqWoYMPe2dSsPiLO9Alf_YIe5zpi-zY4C3aLw5g9at35eZTfNd0gBRpR5ojkMIcZZ6IgAA
// &redirect_uri=https%3A%2F%2Flocalhost%3A12345
// &resource=https%3A%2F%2Fservice.contoso.com%2F
// &client_secret=p@ssw0rd

    var options = {
        uri: 'https://login.microsoftonline.com/common/oauth2/token',
        method: 'POST'
    };
    const formData = {
        code: req.query.code,
        client_id: 'bf411a0c-d417-401d-a2a9-446b63ec8879',
        client_secret: '64Q@mwsL=QqSYKUOEGLDBVU38Cwu@z_8',
        redirect_uri: 'http://localhost:3000/auth/redirect',
        grant_type: 'authorization_code'
    };
    request.post({url: 'https://login.microsoftonline.com/common/oauth2/token', formData: formData}, (error, response, body) => {
        var JSONresponse = JSON.parse(body)
        accessToken = JSONresponse.access_token;
        res.send("Close this dialog and click Approve again")
    })
})

app.post('/action', (req, res) => {
    console.log(req.headers.authorization);
    console.log(accessToken)
    if (!accessToken || accessToken === "") {
        res.status(401).send(JSON.stringify(
            {
                authenticationUrl: "https://login.microsoftonline.com/common/oauth2/authorize?\
                client_id=bf411a0c-d417-401d-a2a9-446b63ec8879\
                &response_type=code\
                &redirect_uri=http%3A%2F%2Flocalhost%3A3000%2Fauth%2Fredirect\
                &response_mode=query\
                &resource=https%3A%2F%2Fgraph.microsoft.com%2F\
                &state=12345"
            }
        )).end()
        console.log("unauthorized")
    }
    else {
        res.status(200).header("CARD-ACTION-STATUS", "Done!").end()
        console.log("success")
    }
})

app.listen(port, () => console.log(`Example app listening on port ${port}!`))