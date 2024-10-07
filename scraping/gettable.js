// Seleccionar todos los elementos de la lista que contienen la información de la empresa
const elements = document.querySelectorAll('li > div > div > div > div > div > div > div');

// Inicializa un array para acumular los resultados
let results = [];

// Iterar sobre todos los elementos seleccionados
elements.forEach(element => {
  // Obtener el nombre de la empresa
  const companyNameElement = element.querySelector('a.ember-view');
  const companyName = companyNameElement ? companyNameElement.textContent.trim() : '';

  // Obtener el otro nombre de la empresa (en muchos casos, puede ser el tipo de industria)
  const otherNameElement = element.querySelector('span:nth-child(1)');
  const otherName = otherNameElement ? otherNameElement.textContent.trim() : '';

  // Obtener la cantidad de empleados en LinkedIn
  const linkedInNameElement = element.querySelector('span:nth-child(3) > a');
  const linkedInName = linkedInNameElement ? linkedInNameElement.textContent.trim() + ' en LinkedIn' : '';

  // Añadir los resultados a la lista solo si al menos uno de los datos no está vacío
  if (companyName || otherName || linkedInName) {
    results.push({
      Empresa: companyName,
      Industria: otherName,
      "Cantidad empleados": linkedInName
    });
  }
});

// Mostrar los resultados en una tabla en la consola
console.table(results);
