const Elements = document.querySelectorAll('li > div > div > div > div > div > div > div > div > span:nth-child(1)');

// Inicializa una variable para acumular los nombres de las empresas
let companyNames = '';

// Itera sobre todos los elementos seleccionados y acumula los nombres no vacíos con saltos de línea
Elements.forEach(element => {
  const companyName = element.textContent.trim();
  if (companyName) {  // Solo agrega nombres no vacíos
    companyNames += companyName + '\n';
  }
});

// Imprime todos los nombres de las empresas, cada uno en una nueva línea
console.log(companyNames.trim());
