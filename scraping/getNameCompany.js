const anchorElements = document.querySelectorAll('li > div > div > div > div > div > div > div > div > div > a.ember-view');

// Inicializa una variable para acumular los nombres de las empresas
let companyNames = '';

// Itera sobre todos los elementos seleccionados y acumula los nombres con saltos de línea
anchorElements.forEach(anchor => {
  companyNames += anchor.textContent.trim() + '\n';
});

// Imprime todos los nombres de las empresas, cada uno en una nueva línea
console.log(companyNames.trim());