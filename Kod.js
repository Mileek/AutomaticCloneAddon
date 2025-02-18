/** 
 * Zmienne Globalne:
 * @param {string} FIRST_WEBSITE_NAME - Nazwa pierwszego serwera ERP, z którego ma być wykonana kopia
 * @param {string} SECOND_WEBSITE_NAME - Nazwa drugiego serwera ERP, na który ma być wykonana kopia
 * @param {string} SUBJECT_TEXT - Tytuł wiadomości po której skrypt będzie miał szukać, to jest MEGA WAŻNE JBC
 * @param {int} DAYS_TO_SEARCH - Wiadomości z ilu dni mają być brane pod uwagę?
 * @param {int} CHECK_EMAIL_INTERVAL - Interwał sprawdzania maili, minimum to 1h (blokada od google)
 */
var FIRST_WEBSITE_NAME = "surix";
var SECOND_WEBSITE_NAME = "Milus";
var SUBJECT_TEXT = "Order clone"; // Tytuł maila do wyszukiwania
var DAYS_TO_SEARCH = 2; // 2 dni do tyłu, czyli szukaj maili z ostatnich 2 dni
var CHECK_EMAIL_INTERVAL = 1; // Minimalny odstęp to 1h

/**
 * Tworzenie UI karty i funkcji dla przycisków Start/Stop
**/
function createHomepage()
{
    var card = CardService.newCardBuilder();

    var section = CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
            .setText("Naciśnij przycisk aby włączyć automatyczne klonowanie co godzinę."))
        .addWidget(CardService.newButtonSet()
            .addButton(CardService.newTextButton()
                .setText("Wyślij i monitoruj")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("startMonitoring"))
                .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
            .addButton(CardService.newTextButton()
                .setText("Przestań monitorować")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("stopMonitoring"))));

    card.addSection(section);
    return card.build();

}

/**
 * Funkcja rozpoczynająca tworzenie komentarzy i podpinająca event do monitorowania skrzynki i wysyłania co godzinę.
**/
function startMonitoring()
{
    // Najpierw usuń wszystkie istniejące triggery
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function (trigger)
    {
        ScriptApp.deleteTrigger(trigger);
    });

    // Teraz utwórz nowy trigger
    createTimeTrigger();
    checkNewEmails(); // Pierwsze uruchomienie od razu

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
            .setText("Rozpoczęto monitorowanie"))
        .build();
}

/**
 * Funkcja usuwająca trigger
**/
function stopMonitoring()
{
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function (trigger)
    {
        return ScriptApp.deleteTrigger(trigger);
    });

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
            .setText("Monitorowanie zatrzymane"))
        .build();
}

/**
 * Tworzenie triggera, który będzie uruchamiany co X czasu, wtedy skrzynka będzie "przeszukiwana"
**/
function createTimeTrigger()
{
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function (trigger)
    {
        return ScriptApp.deleteTrigger(trigger);
    });

    ScriptApp.newTrigger('checkNewEmails')
        .timeBased()
        .everyHours(1) // Minimalny odstęp to 1h
        .create();
}

/**
 * Funkcja pomocnicza, która ułatwi dobór wyszukiwanych maili, jeśli chcemy przeszukać np. 7 dni wstecz
**/
function getSearchDate()
{
    var date = new Date();
    date.setDate(date.getDate() - (DAYS_TO_SEARCH - 1));
    date.setHours(0, 0, 0, 0);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

/**
 * Główna funkcja sprawdzająca maile
 */
function checkNewEmails()
{
    var searchDate = getSearchDate();
    var searchQuery = 'in:anywhere subject:"' + SUBJECT_TEXT + '" after:' + searchDate;

    var threads = GmailApp.search(searchQuery);

    threads.forEach(function (thread)
    {
        const messages = thread.getMessages();
        messages.forEach(function (message)
        {
            if (processEmail(message))
            {
                Logger.log('Order added.');
            }
        });
    });
}

/**
 * Przetwórz pojedynczego maila
 * @return {boolean} true jeśli email został poprawnie przetworzony, w przeciwnym wypadku false
 */
function processEmail(message)
{
    const body = message.getPlainBody();
    const orderId = extractOrderId(body);
    const customerEmail = extractCustomerEmail(body);

    if (!orderId || !customerEmail)
    {
        Logger.log('No order ID or email found in message: ' + message.getId());
        return false;
    }

    var filterDates = getFilterDate();
    var orderData = getOrderDataFromFirstSystem(customerEmail, filterDates[0], filterDates[1]);

    if (!orderData)
    {
        return false;
    }

    const result = createOrderInSecondSystem(orderData);
    return result !== null;
}

/**
 * Funkcja przeszukująca treść maila w celu znalezienia adresu email klienta, !ważne w treściu musi być "Email: " inaczej go nie odnajdzie
 * Nie funkcja nie zamienia ciągu znaków na ToLower ani ToUpper, więc ciąg musi być idealny
 * @param {string} body Treść wiadomości email, jej body, czyli SAMA treśćbez reply to itp, ciągły tekst
**/
function extractCustomerEmail(body)
{
    var match = body.match(/[eE]mail:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    return match ? match[1] : null;
}

/**
 * Wyciągnij ID zamówienia z treści maila, !ważne w treściu musi być "ID: " inaczej go nie odnajdzie
 * Nie funkcja nie zamienia ciągu znaków na ToLower ani ToUpper, więc ciąg musi być idealny
 * @param {string} body Treść wiadomości email, jej body, czyli SAMA treśćbez reply to itp, ciągły tekst
 */
function extractOrderId(body)
{
    const match = body.match(/ID[:\s]*([A-Za-z0-9]+)/i);
    return match ? match[1] : null;
}

/**
 * Pobiera zakres dat do filtrowania zamówień w API Sellasist
 * Obecnie ustawiona data od 1 grudnia 2024 do dnia dzisiejszego * 
 * @returns {Array} Tablica zawierająca:
 *   - dateFrom: Data początkowa w formacie 'YYYY-MM-DD HH:mm:ss'
 *   - dateTo: Data końcowa (dzisiejsza) w formacie 'YYYY-MM-DD HH:mm:ss'
 */
function getFilterDate()
{
    var currentDate = new Date();
    var pastDate = new Date('2024-12-01');

    // Dodaj godzinę do daty końcowej aby uwzględnić strefę czasową
    currentDate.setHours(currentDate.getHours() + 1);

    var dateFrom = pastDate.toISOString().replace('T', ' ').split('.')[0];
    var dateTo = currentDate.toISOString().replace('T', ' ').split('.')[0];

    Logger.log('Date range: ' + dateFrom + ' to ' + dateTo);
    return [dateFrom, dateTo];
}

/**
 * Pobierz dane z pierwszego systemu
 * @param {string} email to email klienta
 */
function getOrderDataFromFirstSystem(email, dateFrom, dateTo)
{
    var url = "https://" + FIRST_WEBSITE_NAME + ".sellasist.pl/api/v1/orders_with_carts" +
        "?offset=0" +
        "&limit=50" +
        "&email=" + encodeURIComponent(email) +
        "&date_from=" + encodeURIComponent(dateFrom) +
        "&date_to=" + encodeURIComponent(dateTo);

    const options = {
        method: 'get',
        headers: {
            'apiKey': getSellasistApiKey("FIRST_API_KEY"),
            'accept': 'application/json'
        },
        muteHttpExceptions: true
    };

    try
    {
        const response = UrlFetchApp.fetch(url, options);
        const orders = JSON.parse(response.getContentText());

        // Sprawdź czy mamy tablicę zamówień i czy nie jest pusta
        if (!Array.isArray(orders) || orders.length === 0)
        {
            Logger.log('No orders found for email: ' + email);
            return null;
        }

        // Znajdź zamówienie z odpowiednim ID
        return orders[0]; // Zwracamy pierwsze (najnowsze) zamówienie
    } catch (err)
    {
        Logger.log('Error getting order from first system: ' + err);
        return null;
    }
}

/**
 * Mapuj zamówienia na dane dla drugiego systemu
 * @param {Object} sourceOrder Zamówienie z pierwszego systemu
 */
function mapOrderData(sourceOrder)
{
    if (!sourceOrder || typeof sourceOrder !== 'object')
    {
        Logger.log('Invalid source order data');
        return null;
    }

    // Kopiujemy całe zamówienie
    const mappedOrder = JSON.parse(JSON.stringify(sourceOrder));

    // Modyfikujemy produkty tylko jeśli to konieczne
    if (Array.isArray(mappedOrder.carts))
    {
        mappedOrder.carts = mappedOrder.carts
            .filter(function (cart)
            {
                return cart && parseFloat(cart.quantity) > 0;
            })
            .map(function (cart)
            {
                // Tworzymy nowy obiekt zachowując oryginalną strukturę
                var mappedCart = {
                    id: cart.id,
                    product_id: cart.product_id,
                    variant_id: cart.variant_id,
                    selected_options: cart.selected_options,
                    name: cart.name,
                    image: cart.image,
                    quantity: cart.quantity,
                    price: cart.price,
                    weight: cart.weight,
                    ean: cart.ean,
                    symbol: cart.symbol,
                    catalog_number: cart.catalog_number,
                    tax_rate: cart.tax_rate,
                    additional_information: cart.additional_information,
                    external_offer_id: cart.external_offer_id
                };
                return mappedCart;
            });
    }

    // Zachowujemy oryginalną strukturę statusu
    mappedOrder.status = sourceOrder.status;

    return mappedOrder;
}

/**
 * Stwórz zamównienie w drugim systemie
 * @param {Object} orderData Dane zamówienia
 */
function createOrderInSecondSystem(orderData)
{
    var url = "https://" + SECOND_WEBSITE_NAME + ".sellasist.pl/api/v1/orders" +
        "?offset=0" +
        "&limit=50";

    const mappedData = mapOrderData(orderData);

    const options = {
        method: 'post',
        headers: {
            'apiKey': getSellasistApiKey("SECOND_API_KEY"),
            'accept': 'application/json',
            'content-type': 'application/json'
        },
        payload: JSON.stringify(mappedData),
        muteHttpExceptions: true
    };
    Logger.log(JSON.stringify(mappedData, null, 2));

    try
    {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseBody = JSON.parse(response.getContentText());

        // Jeśli zamówienie już istnieje, traktujemy to jako sukces, yaay
        if (responseBody.status === "exist")
        {
            Logger.log('Order already exists with ID: ' + responseBody.order_id);
            return responseBody; // Zwracamy odpowiedź, żeby oznaczyć mail jako przetworzony
        }

        if (responseCode !== 200 && responseCode !== 201)
        {
            Logger.log('Error creating order. Status: ' + responseCode);
            return null;
        }

        return responseBody;
    } catch (err)
    {
        Logger.log('Error creating order in second system: ' + err);
        return null;
    }
}

/**
 * Pobierz Api Key z ustawień skryptu
 * @param {string} apiName Nazwa klucza API
 */
function getSellasistApiKey(apiName)
{
    return PropertiesService.getScriptProperties().getProperty(apiName);
}