const inputType = document.getElementById("inputType");

const queryInput = document.getElementById("query");

const numberSection =
    document.getElementById("numberSection");

const statusBox =
    document.getElementById("status");

const downloadSection =
    document.getElementById("downloadSection");

const downloadButton =
    document.getElementById("downloadButton");

const generateButton =
    document.getElementById("generateButton");


inputType.addEventListener("change", function () {

    if (inputType.value === "url") {

        numberSection.style.display = "none";

        queryInput.placeholder =
            "Paste X post URL here";

    } else {

        numberSection.style.display = "block";

        queryInput.placeholder =
            "Example: #AI";
    }

});


async function generateReport() {

    const query = queryInput.value.trim();

    const numberOfPosts =
        document.getElementById(
            "numberOfPosts"
        ).value;


    statusBox.innerHTML = "";

    statusBox.className = "";

    downloadSection.classList.add(
        "hidden"
    );


    if (!query) {

        statusBox.innerHTML =
            "Please enter a keyword, hashtag, or X post URL.";

        statusBox.className = "error";

        return;
    }


    generateButton.disabled = true;

    generateButton.innerText =
        "Generating Report...";


    statusBox.innerHTML =
        "Please wait. Processing X posts...";


    try {

        const response = await fetch(
            "/generate",
            {
                method: "POST",

                headers: {
                    "Content-Type":
                        "application/json"
                },

                body: JSON.stringify({

                    input: query,

                    input_type:
                        inputType.value,

                    number_of_posts:
                        Number(numberOfPosts)

                })
            }
        );


        const data =
            await response.json();


        if (!response.ok) {

            throw new Error(
                "Something went wrong."
            );
        }


        statusBox.innerHTML =
            "Report generated successfully!";

        statusBox.className =
            "success";


        downloadButton.href =
            data.download_url;


        downloadSection.classList.remove(
            "hidden"
        );


    } catch (error) {

        statusBox.innerHTML =
            "Something went wrong.";

        statusBox.className =
            "error";

    }


    generateButton.disabled = false;

    generateButton.innerText =
        "Generate Report";
}