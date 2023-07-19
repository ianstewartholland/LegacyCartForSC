﻿<header class="header navbar-fixed-top">
    <div class="main-header" id="main_header">
        <div class="container">
            <div class="row" id="main_header_row">
                <div class="col-xs-8 col-md-5 addspace">
                    <a href="https://www.chiwater.com/Home" class="nounderline">
                        <img src="https://www.pcswmm.com/Images/topbar_logo.png" class="hidden-sm hidden-xs" />
                        <img src="https://www.pcswmm.com/Images/topbar_logo_only.png" class="hidden-lg hidden-md" />
                    </a>

                    <img src="https://www.pcswmm.com/Images/topbar_divider.png" class="divider" />
                    <a href="cart.asp" class="nounderline" id="site_title_link">
                        <% 
                            strCartItem = " ITEM"
                            If Session("NumCartItems") > 1 Then
                                strCartItem = " ITEMS"
                            End If
                            %>
                        <span class="site-title"><span class="hide-xxs">SHOPPING </span>CART<% If Session("NumCartItems") > 0 Then Response.Write "&nbsp;&nbsp;(" & Session("NumCartItems") & strCartItem & ")" %></span>
                    </a>
                </div>
                <div class="hidden-sm hidden-xs col-md-5">
                </div>
                <div class="col-xs-4 col-md-2 header-icons">
                    <div class="hidden-sm hidden-xs link-container pull-right home-page-links hidden-print" style="margin-top:-5px;">
                        <!--<a href="javascript:void(0)" id="header_resources" name="header_resources" class="nounderline">RESOURCES</a>-->
                        <a href="javascript:void(0)" id="header_resources" name="header_resources" class="btn btn-blue" style="color:white;font-size:17px;background-color:#152831;padding-top: 9px;">
                            CONTINUE BROWSING <img src="https://www.pcswmm.com/images/btn-arrow-down.png" style="padding-left: 10px;vertical-align: middle;"></a>
                    </div>
                    <div class="pull-right hidden-print">
                        <a href="javascript:void(0);" id="menu_link" class="nounderline hidden-lg hidden-md">
                            <img src="https://www.pcswmm.com/images/shared/icon_menu.png" class="header-icon-distance header-icon" alt="Menu" /></a>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div class="clearfix"></div>
    <div class="container hidden hidden-sm hidden-xs" id="resource_menu">
        <div class="row resource-menu">
            <div class="col-xs-6">
                <div class="resource-header"></div>
                <div class="row">
                    <div class="resource-item col-md-4" data-url="https://www.pcswmm.com/">
                        <div class="resource-sub-header">
                            <a href="https://www.pcswmm.com/" class="nounderline">
                                <img src="https://www.pcswmm.com/images/shared/icon_software.png" class="resource-icon img-responsive">
                                Software
                            </a>
                        </div>
                        <div class="row hidden-sm hidden-xs">
                            <div class="col-md-12 resource-text">
                                Tap in to water management modeling that excels. PCSWMM is flexible, easy to use and streamlines your workflow – saving you time and resources.
                            </div>
                        </div>
                    </div>
                    <div class="resource-item col-md-4" data-url="https://www.chiwater.com/Training/">
                        <div class="resource-sub-header">
                            <a href="https://www.chiwater.com/Training/" class="nounderline">
                                <img src="https://www.pcswmm.com/images/shared/icon_training.png" class="resource-icon img-responsive">
                                Training
                            </a>
                        </div>
                        <div class="row hidden-sm hidden-xs">
                            <div class="col-md-12 resource-text">
                                Beginner or seasoned user, our flexible training options help you understand and master the full capabilities of both EPA SWMM5 and PCSWMM.
                            </div>
                        </div>
                    </div>
                    <div class="resource-item col-md-4" data-url="https://www.openswmm.org/">
                        <div class="resource-sub-header">
                            <a href="https://www.openswmm.org/" class="nounderline">
                                <img src="https://www.pcswmm.com/images/shared/icon_community.png" class="resource-icon img-responsive">
                                Community
                            </a>
                        </div>
                        <div class="row hidden-sm hidden-xs">
                            <div class="col-md-12 resource-text">
                                There's a whole community to support you: Open SWMM. Get answers, suggest improvements, share modifications and more with the Knowledge Base and Code Viewer.
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-xs-6 resource-menu-separator">
                <div class="resource-header"></div>
                <div class="row">
                    <div class="resource-item col-md-4" data-url="https://www.chijournal.org/">
                        <div class="resource-sub-header">
                            <a href="https://www.chijournal.org/" class="nounderline">
                                <img src="https://www.pcswmm.com/images/shared/icon_journal.png" class="resource-icon img-responsive">
                                Journal
                            </a>
                        </div>
                        <div class="row hidden-sm hidden-xs">
                            <div class="col-md-12 resource-text">
                                Our peer-reviewed, open-access Journal of Water Management Modeling. Expand your knowledge, get insights and discover new approaches that let you work more effectively.
                            </div>
                        </div>
                    </div>
                    <div class="resource-item col-md-4" data-url="https://www.icwmm.org/">
                        <div class="resource-sub-header">
                            <a href="https://www.icwmm.org/" class="nounderline">
                                <img src="https://www.pcswmm.com/images/shared/icon_conference.png" class="resource-icon img-responsive">
                                Conference
                            </a>
                        </div>
                        <div class="row hidden-sm hidden-xs">
                            <div class="col-md-12 resource-text">
                                The International Conference on Water Management Modeling. Meet your colleagues, share your experiences and be on the forefront of advances in our profession.
                            </div>
                        </div>
                    </div>
                    <div class="resource-item col-md-4" data-url="https://www.chiwater.com/Consulting/">
                        <div class="resource-sub-header">
                            <a href="https://www.chiwater.com/Consulting/" class="nounderline">
                                <img src="https://www.pcswmm.com/images/shared/icon_consulting.png" class="resource-icon img-responsive">
                                Consulting
                            </a>
                        </div>
                        <div class="row hidden-sm hidden-xs">
                            <div class="col-md-12 resource-text">
                                Not sure how to solve a complex water management issue? Put our experience, knowledge, and innovation to work for you.
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div class="container-fluid menu-small hidden-lg hidden-md hidden" id="menu_small">
        <div class="row">
            <div class="col-xs-12">
                <ul>
                    <li><a role="menuitem" tabindex="-1" href="cart.asp">SHOPPING CART</a></li>
                    <li role="presentation" class="divider"></li>
                    <li>
                        <a role="menuitem" tabindex="-1" id="link_resources" class="resources-closed" style="color: #e8eaeb; background-color: #152831;">RESOURCES</a>
                        <div class="row resource-menu resource-menu-small" id="resource_menu_small">
                            <div class="col-xs-5">
                                <div class="resource-header"></div>
                                <div class="row">
                                    <div class="resource-item col-md-4" data-url="https://www.pcswmm.com/">
                                        <div class="resource-sub-header">
                                            <a href="https://www.pcswmm.com/" class="nounderline">
                                                <img src="https://www.pcswmm.com/images/shared/icon_software.png" class="resource-icon img-responsive" />
                                                Software
                                            </a>
                                        </div>
                                    </div>
                                    <div class="resource-item col-md-4" data-url="https://www.chiwater.com/Training/">
                                        <div class="resource-sub-header">
                                            <a href="https://www.chiwater.com/Training/" class="nounderline">
                                                <img src="https://www.pcswmm.com/images/shared/icon_training.png" class="resource-icon img-responsive" />
                                                Training
                                            </a>
                                        </div>
                                    </div>
                                    <div class="resource-item col-md-4" data-url="https://www.openswmm.org/">
                                        <div class="resource-sub-header">
                                            <a href="https://www.openswmm.org/" class="nounderline">
                                                <img src="https://www.pcswmm.com/images/shared/icon_community.png" class="resource-icon img-responsive" />
                                                Community
                                            </a>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="col-xs-7 resource-menu-separator">
                                <div class="resource-header"></div>
                                <div class="row">
                                    <div class="resource-item col-md-4" data-url="https://www.chijournal.org/">
                                        <div class="resource-sub-header">
                                            <a href="https://www.chijournal.org/" class="nounderline">
                                                <img src="https://www.pcswmm.com/images/shared/icon_journal.png" class="resource-icon img-responsive" />
                                                Journal
                                            </a>
                                        </div>
                                    </div>
                                    <div class="resource-item col-md-4" data-url="https://www.icwmm.org/">
                                        <div class="resource-sub-header">
                                            <a href="https://www.icwmm.org/" class="nounderline">
                                                <img src="https://www.pcswmm.com/images/shared/icon_conference.png" class="resource-icon img-responsive" />
                                                Conference
                                            </a>
                                        </div>
                                    </div>
                                    <div class="resource-item col-md-4" data-url="https://www.chiwater.com/Consulting/">
                                        <div class="resource-sub-header">
                                            <a href="https://www.chiwater.com/Consulting/" class="nounderline">
                                                <img src="https://www.pcswmm.com/images/shared/icon_consulting.png" class="resource-icon img-responsive" />
                                                Consulting
                                            </a>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </li>
                </ul>
            </div>
        </div>
    </div>

    <script>
        if (typeof HeaderHelper === 'undefined') {
            StartupService.AddStartupCall(new StartupEntity(function () { HeaderHelper.Initialize(); }));
        } else {
            HeaderHelper.Initialize();
        }
    </script>
    <!-- Global site tag (gtag.js) - Google Analytics -->
    <script async src="https://www.googletagmanager.com/gtag/js?id=UA-8955620-8"></script>
    <script>
        window.dataLayer = window.dataLayer || [];
        function gtag(){dataLayer.push(arguments);}
        gtag('js', new Date());

        gtag('config', 'UA-8955620-8');
    </script>
</header>
