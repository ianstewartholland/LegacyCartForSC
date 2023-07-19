/********************************************************************************************************************************************************************************************************************************************************/
var HeaderHelper = (function () {

    function MoveModalsToBody() {
        $(".modal").detach().appendTo('body');
    }

    return {
        SetBodyContentMargin: function () {
            var $subheader = $("#subheader");
            if ($subheader.length === 1) {
                var $bodyContent = $("#bodycontent");
                $bodyContent.css({ 'margin-top': 60 });
            }
        }, Initialize: function () {
            UpdateDebugValues();
            MoveModalsToBody();
            //OffCanvasMenu.Initialize();
            HeaderHelper.SetBodyContentMargin();
            $(document).on("keyup", function (evt) {
                if (evt.keyCode === 27) {//On press of esc close open menus.
                    ResourceMenuHelper.Hide();
                    SmallScreenMenuHelper.Hide();
                }
            });

            //On click any where on document close open menus.
            $(document).on("click", function (evt) {
                ResourceMenuHelper.Hide();
                SmallScreenMenuHelper.Hide();
            });

            //This is to ensure that document.click doesn't take place.
            $("#resource_menu").on("click", function (evt) {
                evt.stopPropagation();
            });

            //Resources link inside header
            $("#header_resources").on('click', function (evt) {
                $('.dropdown.open').dropdown('toggle');
                ResourceMenuHelper.Toggle();
                evt.stopPropagation();
            });

            //Small screen menu
            $("#menu_link").on('click', function (evt) {
                $('.dropdown.open').dropdown('toggle');
                SmallScreenMenuHelper.Toggle();
                evt.stopPropagation();
            });

            //Edit profile and logoff menu
            $("#user_menu").on('show.bs.dropdown', function () {
                ResourceMenuHelper.Hide();
                SmallScreenMenuHelper.Hide();
            });

            //Small screen resources menu
            $("#link_resources").on("click", function (evt) {
                SmallScreenMenuHelper.ToggleResources();
                evt.stopPropagation();
            });

            $(".resource-item").on("click", function (evt) {
                window.location.href = $(this).data("url");
                evt.stopPropagation();
            });

            //HeaderHelper.SetFixHeader(false);
        }, SetFixHeader: function (blnFixed) {
            //if (blnFixed) {
            //    $("body").removeClass("no-fixed-header");
            //} else {
            //    $("body").addClass("no-fixed-header");
            //}
        }
    };
})();
/********************************************************************************************************************************************************************************************************************************************************/
var SmallScreenMenuHelper = (function () {
    return {
        Hide: function () {
            SmallScreenMenuHelper.Toggle(false);
        }, Toggle: function (show) {
            var $menu_small = $("#menu_small");

            if (show !== true && show !== false) {
                show = $menu_small.hasClass('hidden') === true;
            }

            if (show) {
                $menu_small.removeClass("hidden");
                $menu_small.position({
                    my: 'left top' //"horizontal vertical" 
                    , at: 'left bottom'
                    , of: $('#main_header_row')
                    , using: function (pos, arg) {
                        $menu_small.css({
                            left: 0,
                            top: pos.top + 10
                        });
                    }
                    , collision: "none"
                });
            } else {
                $menu_small.addClass("hidden");
                SmallScreenMenuHelper.ToggleResources(false);
            }
        }, ToggleResources: function (show) {
            var $resource_menu_small = $("#resource_menu_small");

            if (show !== true && show !== false) {
                show = $resource_menu_small.hasClass('hidden') === true;
            }

            var $link_resources = $("#link_resources");

            if (show) {
                $link_resources.addClass("resources-open").removeClass("resources-closed");
                //$resource_menu_small.removeClass("hidden");
            } else {
                $link_resources.addClass("resources-closed").removeClass("resources-open");
                //$resource_menu_small.addClass("hidden");
            }
        }
    };
})();
/********************************************************************************************************************************************************************************************************************************************************/
var ResourceMenuHelper = (function () {
    return {
        Hide: function () {
            ResourceMenuHelper.Toggle(false);
        }, Toggle: function (show) {
            var $resource_menu = $("#resource_menu");
            if ($resource_menu.length !== 1) return;

            if (show !== true && show !== false) {
                show = $resource_menu.hasClass("hidden") === true;
            }

            if (show) {
                $resource_menu.removeClass("hidden");
                $resource_menu.position({
                    my: 'left top' //"horizontal vertical" 
                    , at: 'left bottom'
                    , of: $('#main_header_row')
                    , using: function (pos, arg) {
                        $resource_menu.css({
                            left: pos.left,
                            top: pos.top + 11
                        });
                    }
                    , collision: "none"
                });

            } else {
                if ($resource_menu.is(":visible")) {
                    $resource_menu.addClass("hidden");
                    $resource_menu.css({
                        left: 0,
                        top: 0
                    });
                }
            }

            ChangePointerVisibility(show);

            function ChangePointerVisibility(showPointer) {
                var $resource_pointer = $("#resource_pointer");
                if (showPointer === true) {
                    $("header .pointer").css({
                        'display': 'none'
                    });
                    $resource_pointer.addClass("resource-pointer").removeClass("hidden-pointer");
                } else {
                    $resource_pointer.addClass("hidden-pointer").removeClass("resource-pointer");
                    $("header .pointer").removeAttr('style');
                }
            }
        }
    };
})();
/********************************************************************************************************************************************************************************************************************************************************/

