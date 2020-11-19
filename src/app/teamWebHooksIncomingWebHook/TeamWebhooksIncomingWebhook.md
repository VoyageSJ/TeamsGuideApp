# Exercise - Create incoming webhooks
# https://docs.microsoft.com/en-us/learn/modules/msteams-webhooks-connectors/5-exercise-incoming-webhooks
# 실제 API 생성시 수정하여 사용할것

<!--

const myHeaders = new Headers();
myHeaders.append("Content-Type", "application/json");

const raw = JSON.stringify({
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    "summary": "planet Pluto details",
    "sections": [{
        "activityTitle": "planet Pluto details",
        "activityImage": "https://upload.wikimedia.org/wikipedia/commons/e/ef/Pluto_in_True_Color_-_High-Res.jpg",
        "facts": [{
            "name": "Description",
            "value": "Pluto is an icy planet in the Kuiper belt, a ring of bodies beyond the orbit of Neptune. It was the first Kuiper belt object to be discovered and is the largest known dwarf planet. Pluto was discovered by Clyde Tombaugh in 1930 as the ninth planet from the Sun. After 1992, its status as a planet was questioned following the discovery of several objects of similar size in the Kuiper belt. In 2005, Eris, a dwarf planet in the scattered disc which is 27% more massive than Pluto, was discovered. This led the International Astronomical Union (IAU) to define the term \"planet\" formally in 2006, during their 26th General Assembly. That definition excluded Pluto and reclassified it as a dwarf planet."
        },
        {
            "name": "Order from the sun",
            "value": "9"
        },
        {
            "name": "Known satellites",
            "value": "5"
        },
        {
            "name": "Solar orbit (*Earth years*)",
            "value": "247.9"
        },
        {
            "name": "Average distance from the sun (*km*)",
            "value": "590637500000"
        },
        {
            "name": "Image attribution",
            "value": "NASA/Johns Hopkins University Applied Physics Laboratory/Southwest Research Institute/Alex Parker [Public domain]"
        }]
    }],
    "potentialAction": [{
        "@context": "http://schema.org",
        "@type": "ViewAction",
        "name": "Learn more on Wikipedia",
        "target": ["https://en.wikipedia.org/wiki/Pluto"]
    }]
});


const requestOptions = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow"
};

fetch("https://outlook.office.com/webhook/9971fe01-0663-468d-a8f3-b76c7754466a@bb0178d4-dd75-4fd0-bcbd-25b53b8a3eca/IncomingWebhook/ddbcfe41c52f4ed79133724e130f3a5b/987b30a9-8288-4f65-95f0-68ce0ede897f", requestOptions)
    .then(response => response.text())
    // tslint:disable-next-line: no-console
    .then(result => console.log(result))
    // tslint:disable-next-line: no-console
    .catch(error => console.log("error", error));

-->