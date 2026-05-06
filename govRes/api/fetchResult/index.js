const axios = require('axios');
const cheerio = require('cheerio');
const qs = require('querystring');

module.exports = async function (context, req) {
    const seat = req.query.seat || (req.body && req.body.seat);
    const classType = req.query.class || (req.body && req.body.class);

    if (!seat || !classType) {
        context.res = {
            status: 400,
            body: "Please pass a seat number and class (SSC/HSC) on the query string or in the request body"
        };
        return;
    }

    const initChar = seat.charAt(0).toUpperCase();
    const seatNo = seat.substring(1);
    
    const indexUrl = classType === 'SSC' ? 'https://www.gseb.org/' : 'https://www.gseb.org/Result/Index';
    const postUrl = classType === 'SSC' ? 'https://www.gseb.org/Result/SSCResultView' : 'https://www.gseb.org/Result/ResultView';

    try {
        // 1. GET Index page and capture Cookies
        const indexResponse = await axios.get(indexUrl);
        const cookies = indexResponse.headers['set-cookie'];
        const $ = cheerio.load(indexResponse.data);

        // 2. Extract Security Tokens
        const token = $('input[name="__RequestVerificationToken"]').val();
        const hdnCap = $('#hdnCaptchaAns').val() || $('input[name="hdnCaptcha"]').val();
        const capLabel = $('#lblCaptcha').text();

        // 3. Solve Math Captcha
        const mathMatch = capLabel.replace(/&#x2B;/g, '+').match(/(\d+)\s*[\+\-]\s*(\d+)/);
        if (!mathMatch) throw new Error("Captcha label not found");
        const capAns = (parseInt(mathMatch[1]) + parseInt(mathMatch[2])).toString();

        // 4. POST the data using the SAME session cookies
        const payload = qs.stringify({
            'InitialCharacter': initChar,
            'SeatNo': seatNo,
            'Captcha': capAns,
            'hdnCaptcha': hdnCap,
            '__RequestVerificationToken': token,
            'go': '  Go  ',
            '__Invariant': 'SeatNo'
        });

        const postResponse = await axios.post(postUrl, payload, {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Cookie': cookies ? cookies.join('; ') : ''
            }
        });

        context.res = {
            body: postResponse.data
        };

    } catch (err) {
        context.res = {
            status: 500,
            body: "Proxy Error: " + err.message
        };
    }
};