document.addEventListener('DOMContentLoaded', function () {
    const isUnder18 = document.getElementById('gsignaturepad') !== null; // Check if this is the Under 18 page
    const isTnC = document.getElementById('psignaturepad') !== null; // Check if this is the TnC page
    let formSubmitted = false; // Declare this outside of if blocks

    if (isUnder18) {
        const gcanvas = document.getElementById('gsignaturepad');
        const gsignaturePad = new SignaturePad(gcanvas);

        // Clear signature pad
        document.getElementById('clearunder18').addEventListener('click', function(event) {
            event.preventDefault();
            gsignaturePad.clear();
            console.log('Under 18 Signature cleared');
        });

        document.getElementById('submit').addEventListener('click', function(event) {
            sendsignature(gsignaturePad, 'gsignature', '/under18', event);
        });
    }

    if (isTnC) {
        const pcanvas = document.getElementById('psignaturepad');
        const psignaturePad = new SignaturePad(pcanvas);

        // Clear signature pad
        document.getElementById('clearTnC').addEventListener('click', function(event) {
            event.preventDefault();
            psignaturePad.clear();
            console.log('TnC Signature cleared');
        });

        document.getElementById('submit').addEventListener('click', function(event) {
            sendsignature(psignaturePad, 'psignature', '/tnc', event);
        });
    }

    function sendsignature(sigpad, signatureInput, route, event) {
        event.preventDefault(); // Prevent the default form submission

        if (formSubmitted) {
            return; // Exit if the form has already been submitted
        }

        if (isUnder18) {
            const gname = document.getElementById('gname');
            const gcontact = document.getElementById('gcontact');

            if (gname.isEmpty()) {
                alert("Please provide the Guardian's Name.");
            }

            if (gcontact.isEmpty()) {
                alert("Please provide the Guardian's contact.");
            }
        }

        try {
            if (!sigpad.isEmpty()) {
                // Save as PNG for better quality
                const signatureData = sigpad.toDataURL("image/png");
                document.getElementById(signatureInput).value = signatureData; // Assign to hidden input
                console.log('Signature data URL:', signatureData); // Debugging line
                
                formSubmitted = true; // Set flag to true to prevent further submissions

                const submitButton = document.getElementById('submit');
                submitButton.disabled = true;  // Disable the button
                submitButton.value = 'Submitting...';  // Optionally change the button text

                // Send data via AJAX
                const formData = new FormData();
                formData.append(signatureInput, signatureData); // Append signature data to form

                if (isUnder18) {
                    formData.append('acknowledgement', document.getElementById('acknowledgement').value);
                    formData.append('gname', document.getElementById('gname').value);
                    formData.append('gcontact', document.getElementById('gcontact').value);
                }

                const xhr = new XMLHttpRequest();
                xhr.open('POST', route, true); // Adjust the URL to match your Flask route

                // Handle the AJAX response
                xhr.onload = function () {
                    if (xhr.status === 200) {
                        const response = JSON.parse(xhr.responseText);
                        console.log('Success:', response);  // Log success response
                        window.location.href = response.next_url;  // Redirect to the next page
                    } else {
                        console.error('Error submitting signature:', xhr.statusText);
                        // Optionally handle the error without alerting
                        console.error('Error submitting signature. Status:', xhr.status);
                    }
                };

                xhr.onerror = function () {
                    console.error('Request error');
                    // Optionally handle the error without alerting
                    console.error('Error during submission.');
                };

                xhr.send(formData); // Send the form data via AJAX
            } else {
                alert('Please provide a signature.');
            }
        } catch (error) {
            console.error('Error capturing signature:', error);
            // Log error instead of alerting
            console.error('There was an error capturing your signature. Please try again.');
        }
    }
});
