// Function to load agenda items from emails.json
function loadAgenda() {
    fetch('emails.json')
      .then(response => response.json())
      .then(data => {
        displayAgenda(data);
      })
      .catch(error => {
        console.error('Error loading agenda:', error);
      });
  }
  
  // Function to display agenda items on the page
  function displayAgenda(agendaItems) {
    const agendaList = document.getElementById('agenda-list');
    agendaList.innerHTML = '';
  
    agendaItems.forEach(item => {
      const agendaItem = document.createElement('div');
      agendaItem.className = 'agenda-item';
      agendaItem.innerHTML = `
        <h3>${item.subject}</h3>
        <p>${item.body}</p>
      `;
      agendaList.appendChild(agendaItem);
    });
  }
  
  // Event listener for the Load Agenda button
  document.getElementById('load-agenda').addEventListener('click', loadAgenda);
  
  // Register the service worker for offline capability
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('service-worker.js')
      .then(registration => {
        console.log('Service Worker registered:', registration);
      })
      .catch(error => {
        console.log('Service Worker registration failed:', error);
      });
  }
  