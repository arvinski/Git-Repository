window.addEventListener('DOMContentLoaded', function () {
    var fileInput = document.getElementById('fileInput');
    var fileList = document.getElementById('fileList');

    // Add event listener for file selection
    fileInput.addEventListener('change', function (event) {
        var files = event.target.files;

        // Display selected files
        displayFiles(files);
    });

    // Display selected files
    function displayFiles(files) {
        fileList.innerHTML = '';

        for (var i = 0; i < files.length; i++) {
            var file = files[i];
            var listItem = document.createElement('li');
            listItem.innerHTML = file.name;

            // Add delete button
            var deleteButton = document.createElement('button');
            deleteButton.innerHTML = 'Delete';
            deleteButton.setAttribute('data-file-name', file.name);
            deleteButton.addEventListener('click', function (event) {
                var fileName = event.target.getAttribute('data-file-name');

                // Remove file from list
                removeFile(fileName);
            });

            listItem.appendChild(deleteButton);
            fileList.appendChild(listItem);
        }
    }

    // Remove file from the list
    function removeFile(fileName) {
        var listItem = fileList.querySelector('li[data-file-name="' + fileName + '"]');
        if (listItem) {
            listItem.remove();
        }
    }
});
