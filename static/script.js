// script.js

// Function to set today's date in the date picker
function setTodaysDate() {
    const datePicker = document.getElementById('datePicker');
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0'); // Months are zero-indexed
    const dd = String(today.getDate() - 1).padStart(2, '0');
    const formattedToday = `${yyyy}-${mm}-${dd}`;
    datePicker.value = formattedToday;
}

// Ensure the function runs after the page content has loaded
// document.addEventListener('DOMContentLoaded', setTodaysDate);

// Ensure the function runs after the page content has loaded
document.addEventListener('DOMContentLoaded', function() {
    setTodaysDate();

    // Display alert if error message is present
    if (window.errorMessage) {
        alert(window.errorMessage);
    }
});

// Add event listener to the generate form to include the date in the submission
document.getElementById('generateForm').addEventListener('submit', function(event) {
    const datePicker = document.getElementById('datePicker');
    const dateValue = datePicker.value;
    if (dateValue) {
        const input = document.createElement('input');
        input.type = 'hidden';
        input.name = 'datePicker';
        input.value = dateValue;
        this.appendChild(input);
    }

    // Show the loading message
    document.getElementById('loadingMessage').style.display = 'block';
});

// Add event listener to the upload form to check for file inputs
document.getElementById('uploadForm').addEventListener('submit', function(event) {
    const fileInput = this.querySelector('input[type="file"]');
    if (!fileInput.files.length) {
        alert('Select files to upload!');
        event.preventDefault(); // Prevent form submission
    }
});

if (!!window.EventSource) {
    var source = new EventSource("/status");
    source.onmessage = function(event) {
        console.log('Status update:', event.data);
        if (event.data === 'completed') {
            location.reload();
        } else if (event.data.startsWith('error:')) {
            alert('An error occurred: ' + event.data);
            location.reload();
        }
    };
}
