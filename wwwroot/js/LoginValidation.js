// Validate Login Fields
class OpAuthLogin {
    /*
     * Init Sign In Form Validation in Login Page
     *
     */
    static initValidationLogIn() {
        jQuery('.js-validation-login').validate({
            errorClass: 'invalid-feedback animated fadeInDown',
            errorElement: 'div',
            errorPlacement: (error, e) => {
                jQuery(e).parents('.form-group > div').append(error);
            },
            highlight: e => {
                jQuery(e).closest('.form-group').removeClass('is-invalid').addClass('is-invalid');
            },
            success: e => {
                jQuery(e).closest('.form-group').removeClass('is-invalid');
                jQuery(e).remove();
            },

            rules: {
                'login_username': {
                    required: true,
                    minlength: 3
                },
                'login_password': {
                    required: true,
                    minlength: 8
                }
            },
            messages: {
                'login_username': {
                    required: 'Please enter a username.',
                    minlength: 'Your username must consist of at least 3 characters.'
                },
                'login_password': {
                    required: 'Please provide a password.',
                    minlength: 'Your password must be at least 8 characters long.'
                }
            }
        });

    }

    /*
     * Init functionality
     *
     */
    static init() {
        this.initValidationLogIn();
    }
}

// Initialize when page loads
jQuery(() => { OpAuthLogin.init(); });