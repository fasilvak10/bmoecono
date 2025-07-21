//const API_KEY = PropertiesService.getScriptProperties().getProperty('API_KEY');

/* Gemini Flash 2.0 */
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');


//const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${API_KEY}`
const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`

//generateChat e chat history para mantener la conversación
//system_instruct ion cua es el rol del caracter según selección

//hasta 20 gigas, los archivos se borran cada 2 dias 
//gemini.upload_file(path, display_name)

function geminiAPIChatbot(pregunta, historial) {

    const urls = [
        'https://docs.google.com/document/d/1Ckd_roZkteC0iFs5O92dzBbMB25MJJ9uvtAPcZQKSWw/edit',
        'https://www.argentina.gob.ar/normativa/nacional/ley-27506-324101/actualizacion',
        'https://www.argentina.gob.ar/normativa/nacional/resoluci%C3%B3n-268-2022-376758/actualizacion',
    ];

    const inlineDataArray = [];

    urls.forEach((url) => {
        try {
            if (esGoogleDoc(url)) {
                const docId = extraerDocId(url);
                const texto = leerContenidoDeGoogleDoc(docId);
                const base64Content = Utilities.base64Encode(texto);

                inlineDataArray.push({
                    mimeType: 'text/plain',
                    data: base64Content
                });
            } else {
                const response = UrlFetchApp.fetch(url);
                const headers = response.getHeaders();
                const contentType = headers['Content-Type'] || headers['content-type'] || 'text/plain';
                let fileContent = response.getContentText();

                if (contentType.includes('text')) {
                    fileContent = fileContent.replace(/<[^>]+>/g, ''); // Quitar etiquetas HTML si hay
                }

                const base64Content = Utilities.base64Encode(fileContent);

                inlineDataArray.push({
                    mimeType: 'text/plain',
                    data: base64Content
                });
            }
        } catch (e) {
            Logger.log(`Error procesando ${url}: ${e.message}`);
        }
    });

    const mensajes = historial.map(item => ({
        role: item.sender === "Usuario" ? "USER" : "ASSISTANT",
        parts: [{ text: item.message }]
    }));

    mensajes.push({
        role: "ASSISTANT",
        parts: [
            { text: "Adjunto documentación relevante: Ley 27.506, Resolución 268/2022 y otros documentos útiles." },
            ...inlineDataArray.map(data => ({ inlineData: data }))
        ]
    });

    mensajes.push({
        role: "USER",
        parts: [{ text: pregunta }]
    });

    const payload = {
        contents: mensajes,
        systemInstruction: {
            parts: [{ text: "Responde como si fueras un especialista argentino en derecho y tecnología. Se breve y conciso, y solo extiendete si es muy necesario" }]
        }
    };

    const params = {
        contentType: 'application/json',
        method: 'post',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    const apiResponse = UrlFetchApp.fetch(geminiUrl, params);
    const data = JSON.parse(apiResponse.getContentText());
    Logger.log(JSON.stringify(data, null, 2)); // Para debugging

    return data?.candidates?.[0]?.content?.parts?.[0]?.text || "No se recibió respuesta válida del modelo.";
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

    const instruction = formattedData[1].rol_system_instrution;
    console.log(instruction);

    // Pasar el historial del chat al backend
    const respuestaGemini = geminiAPIChatbot(pregunta, chatHistory);
    return respuestaGemini;
}

function systemInstructionCharacter(character){
    const ss = SpreadsheetApp.openById('14d0kCst55iytMAfC1cMQPRjAOL7n3Hwbc_nfaFAP2qA');
    const characters = ss.getSheetByName('characters');
    const data = characters.getDataRange().getValues();

    const rol_system_instrution = data.filter(c=> c[1] === character).map(i=>{
        return i[2]
    })

    return rol_system_instrution
}

function generarResumenChat(dataFiltrada) {
    if (dataFiltrada) {
        return dataFiltrada.map((item, i) => {
            return `(${i + 1}) ${item["Razón social"]} – CUIT: ${item.Cuit} – Estado: ${item["Estado"]} – Trámite: ${item["Tipo de Trámite"]} – Expte: ${item["Número de expediente"]}`;
        }).join('\n');
    } else {
        return "";
    }
}


function geminiAPI(resultado, dataUser) {
    const datos = generarResumenChat(resultado);
    const role = systemInstructionCharacter(dataUser.character_type);

    let textoPrompt;
    if (datos && datos.trim() !== "") {
        textoPrompt = `Hola Soy ${dataUser.name}. Sé breve, conciso, y natural. Si ${datos} no está vacío, ¿podrías decirme el estado de mis trámites? No más de 1000 caracteres con emijis. No utilices *.`;
    } else {
        textoPrompt = `Hola Soy ${dataUser.name}. Bienvenido al sistema. Por favor hacé click sobre mí para actualizar tus datos.`;
    }

    const payload = {
        contents: [
            {
                role: "USER",
                parts: [{ text: textoPrompt }]
            }
        ],
        systemInstruction: {
            parts: [
                { text: `Actuá como ${role}. Adoptá el estilo de un maestro amable y claro.` }
            ]
        }
    };

    const params = {
        contentType: 'application/json',
        method: 'post',
        payload: JSON.stringify(payload)
    };

    try {
        const response = UrlFetchApp.fetch(geminiUrl, params);
        const data = JSON.parse(response);
        const texto = data?.candidates?.[0]?.content?.parts?.[0]?.text;
        if (!texto) throw new Error("Respuesta vacía");
        console.log(texto);
        return texto;
    } catch (err) {
        console.error("Error al obtener respuesta de Gemini:", err);
        return "⚠️ Hubo un problema al generar la respuesta. Intenta nuevamente.";
    }
}




function leerContenidoDeGoogleDoc(docId) {
    const doc = Docs.Documents.get(docId);
    const bodyElements = doc.body.content;
    let textoPlano = '';

    bodyElements.forEach(elemento => {
        if (elemento.paragraph && elemento.paragraph.elements) {
            elemento.paragraph.elements.forEach(el => {
                if (el.textRun && el.textRun.content) {
                    textoPlano += el.textRun.content;
                }
            });
        }
    });

    return textoPlano;
}

function esGoogleDoc(url) {
    return url.includes('docs.google.com/document');
}

function extraerDocId(url) {
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
}

