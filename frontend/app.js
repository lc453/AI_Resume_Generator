    const analyzeButton =
        document.getElementById("analyze-button");

    const createButton =
        document.getElementById("create-button");

    const downloadButton =
        document.getElementById("download-button");


    const jobDescription =
        document.getElementById("job-description");

    const candidateProfile =
        document.getElementById("candidate-profile");


    const salaryResult =
        document.getElementById("salary-result");

    const fitResult =
        document.getElementById("fit-result");


    const statusMessage =
        document.getElementById("status-message");


    const resumeSection =
        document.getElementById("resume-section");

    const resumePreview =
        document.getElementById("resume-preview");

    const fontSelect =
        document.getElementById("font-select");

    const horizontalMargin =
        document.getElementById("horizontal-margin");

    const verticalMargin =
        document.getElementById("vertical-margin");

    const lineSpacing =
        document.getElementById("line-spacing");

    const embellishment =
        document.getElementById("embellishment");

    const embellishmentValue =
        document.getElementById("embellishment-value");

    const additionalInstructions =
        document.getElementById("additional-instructions");

    const regenerateButton =
        document.getElementById("regenerate-button");

    const rerenderButton =
        document.getElementById("rerender-button");

    const convoContext = [];

    const EmbellishmentStrings = {
        1: "Be careful not to overexaggerate the user's credentials, skills, or experiences at all, keeping only provable facts",
        2: "Be slightly conservative when presenting the user's profile. Don't infer skills unless there is decent evidence to back them up",
        3: "Present the user as a skilled professional, but try not to overstate their skills and experiences",
        4: "Present the user as a skilled proressional and emphasize how their strengths fit the role",
        5: "Strongly emphasize the user's strengths and top experiences, especially as they relate to the job posting.",
        6: "Present the user as confident and capable and use buzzwords that are relevant to the posting. Infer skills based on their profile",
        7: "Present the user as very confident and a perfect fit for the role using powerful buzzwords wherever possible. Round up all numbers as much as justifiable. If a relevant skill is justifiable in any degree by the user's profile, include it",
    };

    /*
        Store the most recently generated DOCX.

        This lets the Download button download
        exactly what was rendered in the preview.
    */

    let currentResume = null
    let currentResumeBlob = null;


    /* ============================================================
    ANALYZE JOB
    ============================================================ */

    analyzeButton.addEventListener(
        "click",
        analyzeJob
    );


    async function analyzeJob() {

        const jobText =
            jobDescription.value.trim();

        const profileText =
            candidateProfile.value.trim();


        if (!jobText) {

            statusMessage.textContent =
                "Please enter a job description.";

            return;
        }


        if (!profileText) {

            statusMessage.textContent =
                "Please enter your background.";

            return;
        }


        disableButtons();

        statusMessage.textContent =
            "Analyzing...";


        salaryResult.textContent = "—";
        fitResult.textContent = "—";


        try {

            const response = await fetch(
                "/api/analyze",
                {
                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body: JSON.stringify({

                        job_description:
                            jobText,

                        candidate_profile:
                            profileText

                    })
                }
            );


            if (!response.ok) {

                throw new Error(
                    "Analysis request failed."
                );

            }


            const result =
                await response.json();


            salaryResult.textContent =
                `$${result.salary_low.toLocaleString()} – $${result.salary_high.toLocaleString()}`;


            fitResult.textContent =
                `${result.fit_score} / 100`;


            statusMessage.textContent =
                "Analysis complete.";

        }

        catch (error) {

            console.error(error);

            statusMessage.textContent =
                "Something went wrong while analyzing the job.";

        }

        finally {

            enableButtons()

        }

    }


    /* ============================================================
    CREATE RESUME
    ============================================================ */

    createButton.addEventListener(
        "click",
        createResume
    );
    regenerateButton.addEventListener(
        "click",
        createResume
    );


    async function createResume() {

        const jobText =
            jobDescription.value.trim();

        const profileText =
            candidateProfile.value.trim();


        if (!jobText) {

            statusMessage.textContent =
                "Please enter a job description.";

            return;
        }


        if (!profileText) {

            statusMessage.textContent =
                "Please enter your background.";

            return;
        }


        disableButtons();


        statusMessage.textContent =
            "Creating resume...";


        try {
            const settings = getResumeSettings();
            if (settings.additional_instructions){
                convoContext.push({"role": "user", "content": settings.additional_instructions})
            }
            const response = await fetch(
                "/api/create-resume",
                {
                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },


                    body: JSON.stringify({

                        job_description: jobText,

                        candidate_profile: profileText,

                        special_instruction: EmbellishmentStrings[settings.embellishment],

                        convo_context: convoContext

                    })
                }
            );


            if (!response.ok) {

                throw new Error(
                    "Resume generation failed."
                );

            }


            /*
                Get the response as binary data.
                const arrayBuffer =
                    await response.arrayBuffer();
            */

            const result = await response.json();
            const convoStep = {"role": "model", "content": result.model_json}
            convoContext.push(convoStep)
            currentResume = result.model_json;

            await renderResume()

            }

        catch (error) {

            console.error(error);

            statusMessage.textContent =
                "Something went wrong while creating the resume.";

        }

        finally {

            enableButtons();

        }

    }

    async function renderResume() {

        if (!currentResume) {
            statusMessage.textContent =
                "No resume info found to render.";
            return;
        }

        statusMessage.textContent =
            "Rendering resume...";

        const settings = getResumeSettings();

        const response = await fetch(
            "/api/render-resume",
            {
                method: "POST",

                headers: {
                    "Content-Type": "application/json"
                },

                body: JSON.stringify({
                    resume_info: currentResume,
                    settings: settings
                })
            }
        );

        if (!response.ok) {
            throw new Error("Resume rendering failed.");
        }

        const arrayBuffer =
            await response.arrayBuffer();

        currentResumeBlob =
            new Blob(
                [arrayBuffer],
                {
                    type:
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                }
            );

        resumeSection.classList.remove("hidden");

        resumePreview.innerHTML = "";

        await docx.renderAsync(
            currentResumeBlob,
            resumePreview,
            null,
            {
                className: "docx",
                inWrapper: true,
                breakPages: true,
                experimental: true,
                renderHeaders: true,
                renderFooters: true,
                renderFootnotes: true,
                renderEndnotes: true
            }
        );

        statusMessage.textContent =
            "Resume created successfully.";

        resumeSection.scrollIntoView({
            behavior: "smooth"
        });
    }
        


    /* ============================================================
    DOWNLOAD RESUME
    ============================================================ */

    downloadButton.addEventListener(
        "click",
        downloadResume
    );


    function downloadResume() {

        if (!currentResumeBlob) {

            return;

        }


        const url =
            URL.createObjectURL(
                currentResumeBlob
            );


        const link =
            document.createElement("a");


        link.href = url;

        link.download =
            "generated_resume.docx";


        document.body.appendChild(link);

        link.click();

        link.remove();


        /*
            Clean up the temporary browser URL.
        */

        URL.revokeObjectURL(url);
    }

    /* ============================================================
    Edit Resume Settings
    ============================================================ */

    embellishment.addEventListener(
        "input",
        updateEmbellishment
    );


    function updateEmbellishment() {

        embellishmentValue.textContent =
            `${embellishment.value} / 7`;

    }

    function getResumeSettings() {

        return {
            font: fontSelect.value,

            horizontal_margin:
                parseFloat(
                    horizontalMargin.value
                ),

            vertical_margin:
                parseFloat(
                    verticalMargin.value
                ),

            line_spacing:
                parseFloat(
                    lineSpacing.value
                ),

            embellishment:
                parseInt(
                    embellishment.value,
                    10
                ),

            additional_instructions:
                additionalInstructions.value.trim()
        };
    }

    rerenderButton.addEventListener(
        "click",
        renderResume
    );

    function disableButtons() {
        analyzeButton.disabled = true;
        createButton.disabled = true;
        regenerateButton.disabled = true;
        rerenderButton.disabled = true;
    }

    function enableButtons() {
        analyzeButton.disabled = false;
        createButton.disabled = false;
        regenerateButton.disabled = false;
        rerenderButton.disabled = false;
    }
