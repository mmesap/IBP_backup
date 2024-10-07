const anchorElements = document.querySelectorAll('li > div > div > div > div > div > div > div > div > span:nth-child(3) > a');

// Inicializa una variable para acumular los nombres de las empresas
let companyNames = '';

// Itera sobre todos los elementos seleccionados y acumula los nombres no vacíos con saltos de línea
anchorElements.forEach(anchor => {
  const companyName = anchor.textContent.trim();
  if (companyName) {  // Solo agrega nombres no vacíos
    companyNames += companyName + ' en LinkedIn'+ '\n';
  }
});

// Imprime todos los nombres de las empresas, cada uno en una nueva línea
console.log(companyNames.trim());