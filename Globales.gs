// Funciona... solo hay que agregaro al spreadsheet
// Deberia generarme un library de esto...

// Source: https://www.youtube.com/watch?v=vdP6sZKp4hU

/*
VLOOKUP(valor_búsqueda; intervalo; índice; [está_ordenado])
VLOOKUP(10003; A2:B26; 2; FALSE)

INFORMACIÓN
Búsqueda vertical. Busca un valor en la primera columna de un intervalo y ofrece el valor de una celda específica en la fila encontrada.
valor_búsqueda>>> Valor que se va a buscar. Por ejemplo, 42, "Gatos" o I24.
intervalo>> Intervalo de la búsqueda. El valor especificado en el argumento valor_búsqueda se busca en la primera columna del intervalo.
índice
Índice de columna del valor que se va a ofrecer, donde a la primera columna de intervalo se le asigna el número 1.
está_ordenado - [opcional]
Indica si la primera columna en la que se va a buscar (la primera columna del intervalo especificado) está ordenado. En tal caso, se ofrece la coincidencia más exacta con el argumento valor_búsqueda.
*/

// Global
const hojaSource = '1-ST22'
const hojaSearch = '1-SEARCH_IN'