<?php

/**
 * Based roughly on wp-login.php @revision 19414
 * http://core.trac.wordpress.org/browser/trunk/wp-login.php?rev=19414
 */

global $wp_version, $Password_Protected, $error, $is_iphone;

/**
 * WP Shake JS
 */
if ( ! function_exists( 'wp_shake_js' ) ) {
	function wp_shake_js() {
		global $is_iphone;
		if ( $is_iphone ) {
			return;
		}
		?>
        <script>
            addLoadEvent = function (func) {
                if (typeof jQuery != "undefined") jQuery(document).ready(func); else if (typeof wpOnload != 'function') {
                    wpOnload = func;
                } else {
                    var oldonload = wpOnload;
                    wpOnload = function () {
                        oldonload();
                        func();
                    }
                }
            };

            function s(id, pos) {
                g(id).left = pos + 'px';
            }

            function g(id) {
                return document.getElementById(id).style;
            }

            function shake(id, a, d) {
                c = a.shift();
                s(id, c);
                if (a.length > 0) {
                    setTimeout(function () {
                        shake(id, a, d);
                    }, d);
                } else {
                    try {
                        g(id).position = 'static';
                        wp_attempt_focus();
                    } catch (e) {
                    }
                }
            }

            addLoadEvent(function () {
                var p = [15, 30, 15, 0, -15, -30, -15, 0];
                p = p.concat(p.concat(p));
                var i = document.forms[0].id;
                g(i).position = 'relative';
                shake(i, p, 20);
            });
        </script>
        <?php
	}
}

/**
 * @since 3.7.0
 */
if ( ! function_exists( 'wp_login_viewport_meta' ) ) {
	function wp_login_viewport_meta() {
		?>
		<meta name="viewport" content="width=device-width" />
		<?php
	}
}

nocache_headers();
header( 'Content-Type: ' . get_bloginfo( 'html_type' ) . '; charset=' . get_bloginfo( 'charset' ) );

// Set a cookie now to see if they are supported by the browser.
setcookie( TEST_COOKIE, 'WP Cookie check', 0, COOKIEPATH, COOKIE_DOMAIN );
if ( SITECOOKIEPATH != COOKIEPATH ) {
	setcookie( TEST_COOKIE, 'WP Cookie check', 0, SITECOOKIEPATH, COOKIE_DOMAIN );
}

// If cookies are disabled we can't log in even with a valid password.
if ( isset( $_POST['password_protected_cookie_test'] ) && empty( $_COOKIE[ TEST_COOKIE ] ) ) {
	$Password_Protected->errors->add( 'test_cookie', __( "<strong>ERROR</strong>: Cookies are blocked or not supported by your browser. You must <a href='http://www.google.com/cookies.html'>enable cookies</a> to use WordPress.", 'password-protected' ) );
}

// Shake it!
$shake_error_codes = array( 'empty_password', 'incorrect_password' );
if ( $Password_Protected->errors->get_error_code() && in_array( $Password_Protected->errors->get_error_code(), $shake_error_codes ) ) {
	add_action( 'password_protected_login_head', 'wp_shake_js', 12 );
}

// Obey privacy setting
add_action( 'password_protected_login_head', 'wp_no_robots' );

add_action( 'password_protected_login_head', 'wp_login_viewport_meta' );

?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" <?php language_attributes(); ?>>
<head>

<meta http-equiv="Content-Type" content="<?php bloginfo( 'html_type' ); ?>; charset=<?php bloginfo( 'charset' ); ?>" />
<title><?php echo apply_filters( 'password_protected_wp_title', get_bloginfo( 'name' ) ); ?></title>

<?php

if ( version_compare( $wp_version, '3.9-dev', '>=' ) ) {
	wp_admin_css( 'login', true );
} else {
	wp_admin_css( 'wp-admin', true );
	wp_admin_css( 'colors-fresh', true );
}

?>

<style media="screen">
#login_error, .login .message, #loginform { margin-bottom: 20px; }
</style>

<?php

if ( $is_iphone ) {
	?>
	<meta name="viewport" content="width=320; initial-scale=0.9; maximum-scale=1.0; user-scalable=0;" />
	<style media="screen">
	.login form, .login .message, #login_error { margin-left: 0px; }
	.login #nav, .login #backtoblog { margin-left: 8px; }
	.login h1 a { width: auto; }
	#login { padding: 20px 0; }
	</style>
	<?php
}

do_action('login_enqueue_scripts');
do_action('password_protected_login_head');

?>

</head>
<body class="login login-password-protected login-action-password-protected-login wp-core-ui">

<div id="login">
    <h1>
        <a style="background-image: url(https://exetercontemporarychoir.com/wp-content/uploads/2020/08/Contemporary-Choir-Logo.jpeg)"
           href="https://exetercontemporarychoir.com"
           title="<?php echo esc_attr(apply_filters('password_protected_login_headertitle', get_bloginfo('name'))); ?>"><?php bloginfo('name'); ?></a>
    </h1>

    <?php do_action('password_protected_login_messages'); ?>
    <p style="text-align: center"><?php _e("This is site is for members of Contemporary Choir only.") ?><br>
        <?php _e("You can find the password in our latest weekly email.") ?><br>
        <?php _e("If you think that you should be able to access this site but can't, please do ") ?><a target="_blank"
                                                                                                        href="https://exetercontemporarychoir.com/stay-in-touch/"><?php _e("get in touch") ?></a><?php _e(" with us.") ?>
    </p>
    <?php do_action('password_protected_before_login_form'); ?>

    <form name="loginform" id="loginform" action="<?php echo esc_url($Password_Protected->login_url()); ?>"
          method="post">
        <p>
            <label for="password_protected_pass"><?php echo apply_filters('password_protected_login_password_title', __('Password', 'password-protected')); ?>
                <br/>
                <input type="password" name="password_protected_pwd" id="password_protected_pass" class="input" value=""
                       size="20" tabindex="20"/></label>
        </p>

        <?php do_action('anr_captcha_form_field') ?><br>

        <?php if ($Password_Protected->allow_remember_me()) : ?>
            <p class="forgetmenot">
                <label for="password_protected_rememberme"><input name="password_protected_rememberme" type="checkbox"
                                                                  id="password_protected_rememberme" value="1"
                                                                  tabindex="90"/> <?php esc_attr_e('Remember Me'); ?>
                </label>
            </p>
        <?php endif; ?>

        <p class="submit">
            <input type="submit" name="wp-submit" id="wp-submit" class="button button-primary button-large"
                   value="<?php esc_attr_e('Log In'); ?>" tabindex="100"/>
            <input type="hidden" name="password_protected_cookie_test" value="1"/>
			<input type="hidden" name="password-protected" value="login" />
			<input type="hidden" name="redirect_to" value="<?php echo esc_attr( $_REQUEST['redirect_to'] ); ?>" />
		</p>
	</form>

	<?php do_action( 'password_protected_after_login_form' ); ?>

</div>

<script>
try{document.getElementById('password_protected_pass').focus();}catch(e){}
if(typeof wpOnload=='function')wpOnload();
</script>

<?php do_action( 'login_footer' ); ?>

<div class="clear"></div>

</body>
</html>
