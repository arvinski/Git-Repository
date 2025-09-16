class DtInitDatatables {
    /*
     * Override a few DataTable defaults
     *
     */
    static exDataTable() {
        jQuery.extend(jQuery.fn.dataTable.ext.classes, {
            sWrapper: "dataTables_wrapper dt-bootstrap4"
        });
    }

    /*
     * Init User DataTable
     * Admin/Index
     */
    static initUserDataTable() {
        jQuery('.js-usr-dataTable').dataTable({            
            columnDefs: [{ orderable: false, targets: [7] }],
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false
        });
    }

    /*
     * Init Service Ticket DataTable
     * Home/Index
     */
    static initTicketDataTable() {
        jQuery('.js-service-dataTable').dataTable({
            pageLength: 5,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false,
            columnDefs: [
                { 'width': '40%', 'targets': 1 },
                { 'className': 'dt-wrap', 'targets': 1 }
            ]
        });
    }

    /*
     * Init CM Activity DataTable
     * Activity/CustomerService
     */
    static initCMActivityDataTable() {
        jQuery('.js-cmactivity-dataTable').dataTable({
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false,
            order: [[3, 'desc']]
        });
    }

    /*
     * Init CM Activity DataTable
     * Activity/CustomerService
     */
    static initSearchICBADataTable() {
        jQuery('.js-searchicba-dataTable').dataTable({
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false
        });
    }

    /*
     * Init ATM DataTable
     * Admin/Payment Institution, etc.
     */
    static initATMDataTable() {
        jQuery('.js-atm-dataTable').dataTable({
            columnDefs: [{ orderable: false, targets: [1] }],
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false
        });
    }

    /*
     * Init ATM DataTable
     * Admin/Payment Institution, etc.
     */
    static initATMLocDataTable() {
        jQuery('.js-atmloc-dataTable').dataTable({
            columnDefs: [{ orderable: false, targets: [2] }],
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false
        });
    }

    static initATMLogsDataTable() {
        jQuery('.js-atmlogs-dataTable').dataTable({
            order: [[3, 'desc']],
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false
        });
    }

    /*
 * Init CM Activity DataTable
 * Activity/CustomerService
 */
    static initATMActivityDataTable() {
        jQuery('.js-atmactivity-dataTable').dataTable({
            order: [[3, 'desc']],
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false,
            searching: false
        });
    }

    static initUserMenu() {
        jQuery('.js-usr-menu').dataTable({
            columnDefs: [{ orderable: false, targets: [3] }],
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false
        });
    }

    static initUserRight() {
        jQuery('.js-usr-rights').dataTable({
            columnDefs: [{ orderable: false, targets: [2] }],
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false
        });
    }

    static initReasonCodes() {
        jQuery('.js-reason_code').dataTable({
            columnDefs: [{ orderable: false, targets: [8] }],
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false
        });
    }

    static initDisableSearch() {
        jQuery('.js-disable').dataTable({
            lengthChange: false,
            paging: false,
            searching: false,
            info: false
        });
    }

    static initCMActivityDataTable() {
        jQuery('.js-cmactivity').dataTable({
            order: [[3, 'desc']],
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false,
            searching: false
        });
    }

    static initCMEntryDataTable() {
        jQuery('.js-cmentry').dataTable({
            pageLength: 8,
            lengthMenu: [[5, 8, 15, 20], [5, 8, 15, 20]],
            autoWidth: false
        });
    }

    /*
     * Init functionality
     *
     */
    static init() {
        this.exDataTable();
        this.initUserDataTable();
        this.initTicketDataTable();
        this.initCMActivityDataTable();
        this.initSearchICBADataTable();
        this.initATMDataTable();
        this.initATMLocDataTable();
        this.initATMLogsDataTable();
        this.initATMActivityDataTable();
        this.initUserMenu();
        this.initUserRight();
        this.initReasonCodes();
        this.initDisableSearch();
        this.initCMEntryDataTable();
    }
}

// Initialize when page loads
jQuery(() => { DtInitDatatables.init(); });