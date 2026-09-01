const queryInput =
    document.getElementById("query");

const numberOfPostsInput =
    document.getElementById(
        "numberOfPosts"
    );

const statusBox =
    document.getElementById("status");

const downloadSection =
    document.getElementById(
        "downloadSection"
    );

const downloadButton =
    document.getElementById(
        "downloadButton"
    );

const generateButton =
    document.getElementById(
        "generateButton"
    );


async function generateReport() {

    const query =
        queryInput.value.trim();

    const numberOfPosts =
        Number(
            numberOfPostsInput.value
        );


    statusBox.innerHTML = "";

    statusBox.className = "";

    downloadSection.classList.add(
        "hidden"
    );


    if (!query) {

        statusBox.innerHTML =
            "Please enter a keyword or hashtag.";

        statusBox.className =
            "error";

        return;
    }


    if (
        numberOfPosts < 1 ||
        numberOfPosts > 100
    ) {

        statusBox.innerHTML =
            "Please enter between 1 and 100 posts.";

        statusBox.className =
            "error";

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

                    input_type: "search",

                    number_of_posts:
                        numberOfPosts

                })
            }
        );


        const data =
            await response.json();


        if (!response.ok) {

            throw new Error(
                data.error ||
                "Something went wrong."
            );
        }


        statusBox.innerHTML =
            `Report generated successfully! ${data.posts} post(s) captured.`;

        statusBox.className =
            "success";


        downloadButton.href =
            data.download_url;


        downloadSection.classList.remove(
            "hidden"
        );


    } catch (error) {

        statusBox.innerHTML =
            error.message ||
            "Something went wrong.";

        statusBox.className =
            "error";

    }


    generateButton.disabled = false;

    generateButton.innerText =
        "Generate Report";
}