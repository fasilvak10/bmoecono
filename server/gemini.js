/* Gemini Flash 1.5 */
const API_KEY = 'AIzaSyA3VkXui9bE-i-NDsrhdYdNn0vP2J709G0';

/* Gemini Flash 1.5 */
const GEMINI_API_KEY = 'AIzaSyCCHdmdbRu4OQ3ZA2AUK8k5dIwcNttrKrI';

//const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${API_KEY}`
const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`

//generateChat e chat history para mantener la conversación
//system_instruction cua es el rol del caracter según selección

//hasta 20 gigas, los archivos se borran cada 2 dias 
//gemini.upload_file(path, display_name)

const datos = `Analista Asignado (evaluador),En Revisión,Informe Técnico Firmado,No corresponde,Providencia Notificada,Subsanación
Brenda,,1,7,58,
Cintia,1,,1,21,
Claudia,2,1,6,71,
Constanza,,,1,4,
Elias,1,3,3,53,1
Facundo O,3,2,7,83,
Florencia,,,9,73,
Franca,,,4,47,
Franco,,,1,12,
Jesica,,,1,5,
Julieta F,,,,5,
Luana,,,,4,
MarianoF,1,,8,68,1
Melanie,,,1,17,
Romina,,,3,11,
Suma total,8,7,52,532,2`;

function geminiAPIChatbot(pregunta, historial) {

    const urls = [
        'https://www.argentina.gob.ar/normativa/nacional/ley-27506-324101/actualizacion',
        'https://www.argentina.gob.ar/normativa/nacional/resoluci%C3%B3n-268-2022-376758/actualizacion',
        'https://lookerstudio.google.com/reporting/b60752fd-4a99-48a3-9b6a-f32d20292b5c'
    ];

    // Inicializar un array para almacenar los contenidos de las URLs en Base64
    const inlineDataArray = [];

    // Iterar sobre cada URL, obtener su contenido y codificarlo en Base64
    urls.forEach((url) => {
        const response = UrlFetchApp.fetch(url);
        const fileContent = response.getContentText().replace(/<[^>]+>/g, ''); // Elimina etiquetas HTML
        const base64Content = Utilities.base64Encode(fileContent);

        inlineDataArray.push({
            "mimeType": "text/html",
            "data": base64Content
        });
    });

    const mensajes = historial.map(item => ({
        role: item.sender === "Usuario" ? "USER" : "ASSISTANT",
        parts: [
            { text: item.message },
            ...(item.sender === "ASSISTANT"
                ? inlineDataArray.map(data => ({
                    "inlineData": data
                }))
                : [])
        ]
    }));

    // Agregar el mensaje actual del usuario y el contenido codificado al final del historial
    mensajes.push({
        role: "USER",
        parts: [
            { text: pregunta },
            ...inlineDataArray.map(data => ({
                inlineData: data
            }))
        ]
    });

    // Crear el payload
    const payload = {
        "contents": mensajes,
        "systemInstruction": {
            "parts": [
                { "text": "Responde como un especialista en temas relaciones a Economía del Conocimiento. Responde en no más de 3000 caracteres. Utiliza emojis, pero pocos, sin abrumar" },
            ]
        }
    };

    // Configurar los parámetros de la solicitud
    const params = {
        'contentType': 'application/json',
        'method': 'post',
        'payload': JSON.stringify(payload)
    };

    // Realizar la solicitud
    const apiResponse = UrlFetchApp.fetch(geminiUrl, params);
    const data = JSON.parse(apiResponse.getContentText());

    // Retornar la respuesta procesada
    return data.candidates[0].content.parts[0].text;
}

function getRespuestaGemini(text, chatHistory) {

    const pregunta = text;
    const ss = SpreadsheetApp.openById('14d0kCst55iytMAfC1cMQPRjAOL7n3Hwbc_nfaFAP2qA');
    const characters = ss.getSheetByName('characters');
    const data = characters.getDataRange().getValues();

    // Obtiene los encabezados (la primera fila)
    const headers = data[0];

    const formattedData = data.slice(1).map(row => {
        let obj = {};
        row.forEach((cell, index) => {
            obj[headers[index]] = cell;
        });
        return obj;
    });

    const instruction = formattedData[2].rol_system_instrution;
    console.log(instruction);

    // Pasar el historial del chat al backend
    const respuestaGemini = geminiAPIChatbot(pregunta, chatHistory);
    return respuestaGemini;
}

function geminiAPI() {
    const payload = {
        "contents": [
            {
                "role": "USER",
                "parts": [
                    {
                        "text": `Considerando el número de tramites "sin evaluar" o de en estado de evaluación o subsanación, no providencia notificada o no corresponde. Sé breve, conciso, y natural. Quiero que respondas como si fueras un argentino rioplatense, sobre el usuario. ${datos}. Soy Brenda, ¿podrías decirme el estado de mis trámites? No más de 500 caracteres con emijis. No utilices *`
                    }
                ]
            }
        ],
        "systemInstruction": {
            //"role": string,
            "parts": [
                {
                    "text": "responde como un maestro"
                }
            ]
        }
    };


    const params = {
        'contentType': 'application/json',
        'methd': 'post',
        'payload': JSON.stringify(payload)
    }

    const response = UrlFetchApp.fetch(geminiUrl, params);
    const data = JSON.parse(response);
    console.log(data.candidates[0].content.parts[0].text)
    return data.candidates[0].content.parts[0].text;
}
