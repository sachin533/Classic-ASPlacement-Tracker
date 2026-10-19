/* ==========================================================================
   Placement Tracker - static/app.js
   Two small, dependency-free enhancements:

     1. the clear (x) control that appears in every search box, and
     2. the delete confirmation dialog.

   Everything here is progressive enhancement - without JavaScript the search
   boxes still submit normally (and the browser's own clear affordance is
   available), the delete forms fall back to the browser's confirm() dialog,
   and the server always re-checks the id, the CSRF token and every field.
   ========================================================================== */

/* --- 1. clear-search control ------------------------------------------- */
(function () {
  'use strict';

  var inputs = document.querySelectorAll('.js-search-input');
  var i, input, form, clearBtn;

  function wire(field) {
    form = field.form;
    clearBtn = form ? form.querySelector('.js-search-clear') : null;
    if (!clearBtn) { return; }

    /* The value the server rendered: only a search that was actually
       submitted should be cleared by submitting an empty query. */
    var submitted = field.value;

    function sync() {
      if (field.value) { clearBtn.removeAttribute('hidden'); }
      else { clearBtn.setAttribute('hidden', 'hidden'); }
    }

    field.addEventListener('input', sync, false);
    field.addEventListener('search', sync, false);
    sync();

    clearBtn.addEventListener('click', function () {
      field.value = '';
      sync();
      field.focus();
      /* Leaving a submitted search on screen after clicking x would be a
         lie - hand back the full, page-1 result set right away. */
      if (submitted !== '' && form) { form.submit(); }
    }, false);
  }

  for (i = 0; i < inputs.length; i++) {
    input = inputs[i];
    wire(input);
  }

  /* The compact bar search has no js-search-input class marker on every
     page, so wire any remaining clear buttons by their own field too. */
  var loose = document.querySelectorAll('.site-search .js-search-clear');
  for (i = 0; i < loose.length; i++) {
    var btn = loose[i];
    var host = btn.parentNode;
    var field = host ? host.querySelector('input[type="search"]') : null;
    if (field && !field.classList.contains('js-search-input')) {
      field.classList.add('js-search-input');
      wire(field);
    }
  }
}());

/* --- 2. delete confirmation dialog -------------------------------------- */
(function () {
  'use strict';

  var overlay = document.getElementById('delete-modal');
  if (!overlay) { return; }

  var confirmBtn = document.getElementById('delete-modal-confirm');
  var lastFocused = null;
  var pendingForm = null;

  var FOCUSABLE = 'a[href], button:not([disabled]), input:not([disabled]), select, textarea, [tabindex]:not([tabindex="-1"])';

  function focusable() {
    return Array.prototype.slice.call(overlay.querySelectorAll(FOCUSABLE))
      .filter(function (el) { return el.offsetParent !== null; });
  }

  function open(form) {
    pendingForm = form;
    lastFocused = document.activeElement;
    overlay.hidden = false;
    document.body.style.overflow = 'hidden';
    var first = focusable()[0];
    if (first) { first.focus(); }
  }

  function close() {
    overlay.hidden = true;
    document.body.style.overflow = '';
    pendingForm = null;
    if (lastFocused && document.contains(lastFocused)) {
      lastFocused.focus();
    }
    lastFocused = null;
  }

  function isOpen() { return !overlay.hidden; }

  /* Intercept the submit of any form that posts to delete.asp. */
  document.addEventListener('click', function (event) {
    var trigger = event.target.closest ? event.target.closest('.js-delete-trigger') : null;
    if (!trigger) { return; }

    var form = trigger.form || document.getElementById(trigger.getAttribute('form'));
    if (!form || form.method.toLowerCase() !== 'post') { return; }

    /* Let the browser confirm() run when scripting is unavailable. */
    event.preventDefault();
    open(form);
  }, false);

  document.addEventListener('submit', function (event) {
    var form = event.target;
    if (!form.classList || !form.classList.contains('js-delete-form')) { return; }
    if (form.getAttribute('data-confirmed') === 'true') { return; }
    /* Keyboard submit (Enter in a field) also gets the dialog. */
    event.preventDefault();
    open(form);
  }, false);

  if (confirmBtn) {
    confirmBtn.addEventListener('click', function () {
      var form = pendingForm;
      if (!form) { close(); return; }
      form.setAttribute('data-confirmed', 'true');
      confirmBtn.disabled = true;
      confirmBtn.textContent = 'Deleting...';
      /* form.submit() bypasses onsubmit, which is what we want here:
         the dialog has already been answered. */
      if (typeof form.requestSubmit === 'function') { form.requestSubmit(); }
      else { form.submit(); }
    }, false);
  }

  overlay.addEventListener('click', function (event) {
    if (event.target === overlay) { close(); return; }
    if (event.target.closest && event.target.closest('[data-modal-close]')) { close(); }
  }, false);

  document.addEventListener('keydown', function (event) {
    if (!isOpen()) { return; }

    if (event.key === 'Escape' || event.keyCode === 27) {
      event.preventDefault();
      close();
      return;
    }

    if (event.key === 'Tab' || event.keyCode === 9) {
      var items = focusable();
      if (!items.length) { return; }
      var first = items[0];
      var last = items[items.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    }
  }, false);
}());
