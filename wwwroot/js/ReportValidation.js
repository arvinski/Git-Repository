// Validate Report Fields
class OpAuthReport {
    /*
     * Init Sign In Form Validation in Report Page
     *
     */
    static initValidationReport() {
        jQuery('.js-validation-report').validate({
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
                'ReportType': {
                    required: true
                }
            },

            messages: {
                'ReportType': {
                    required: 'Please select report type'
                    
                }
            }
        });

    }

    /*
     * Init functionality
     *
     */
    static init() {
        this.initValidationReport();
    }
}

// Initialize when page loads
jQuery(() => { OpAuthReport.init(); });