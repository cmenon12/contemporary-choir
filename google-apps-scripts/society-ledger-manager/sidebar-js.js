function formSuccessHandler(result) {
  if (result == null) {
    result = "success!!"
  }
  document.getElementById("result").innerHTML = `<b>${result}</b>`
}


function formFailureHandler(error) {
  document.getElementById("result").className = "error"
  document.getElementById("result").innerHTML = `<b>${error.message}</b>`
}


/**
 * Checks that the form was filled out correctly.
 */
function checkForm() {

  // Track if any errors occurred
  let success = true;

  // Check the group ID, email, and password
  const formFieldIds = ["groupId", "email", "password"];
  for (let i = 0; i < formFieldIds.length; i++) {
    console.log(formFieldIds[i])
    if (document.getElementById(formFieldIds[i]).value === "") {
      document.getElementById(`${formFieldIds[i]}Error`).style.display = "inline";
      success = false;
    } else {
      document.getElementById(`${formFieldIds[i]}Error`).style.display = "none";
    }
  }

  // Check that they chose at least one action
  if (document.getElementById("savePDF").checked === false &&
    document.getElementById("saveXLSX").checked === false &&
    document.getElementById("import").checked === false) {
    document.getElementById("choseNothingError").style.display = "inline";
    success = false;
  } else {
    document.getElementById("choseNothingError").style.display = "none";
  }

  // Return the result
  return success;

}


/**
 * Runs when the user attempts to submit the form.
 * Processes the form before actually submitting it.
 */
function processForm(action) {

  // Stop now if the form validation failed
  if (action !== "Clear saved data" && !checkForm()) {
    return;
  }

  // Actually submit it
  document.getElementById("ledgerForm").submit();

  // Update the result text
  document.getElementById("actionInput").value = action;
  if (action === "Clear saved data") {
    document.getElementById("result").innerHTML = "Clearing..."
  } else if (action === "GO!") {
    document.getElementById("result").innerHTML = "Downloading..."
  }
}


/**
 * Runs when the form is actually submitted using the
 * processForm(action) function above.
 */
function handleFormSubmit(form) {
  google.script.run
    .withSuccessHandler(formSuccessHandler)
    .withFailureHandler(formFailureHandler)
    .processSidebarForm(form);
}


/**
 * Prevents the form submitting when the sidebar loads.
 */
function preventFormSubmit() {
  const forms = document.querySelectorAll("form");
  for (let i = 0; i < forms.length; i++) {
    forms[i].addEventListener("submit", function (event) {
      event.preventDefault();
    });
  }
}

window.addEventListener("load", preventFormSubmit);
