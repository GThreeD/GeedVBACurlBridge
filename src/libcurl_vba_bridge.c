#include <curl/curl.h>
#include <stdint.h>
#include <stdlib.h>
#include <string.h>

#ifdef __cplusplus
extern "C" {
#endif

#ifndef WSC_BRIDGE_API_VERSION
#define WSC_BRIDGE_API_VERSION 1
#endif

#ifndef WSC_BRIDGE_VERSION_STRING
#define WSC_BRIDGE_VERSION_STRING "0.0.0.0"
#endif

long wsc_bridge_api_version(void) {
    return WSC_BRIDGE_API_VERSION;
}

const char *wsc_bridge_version_string(void) {
    return WSC_BRIDGE_VERSION_STRING;
}

const char *wsc_bridge_name(void) {
    return "libcurl_vba_bridge";
}

long wsc_global_init(void) {
    return (long)curl_global_init(CURL_GLOBAL_DEFAULT);
}

void wsc_global_cleanup(void) {
    curl_global_cleanup();
}

const char *wsc_libcurl_version(void) {
    return curl_version();
}

static void wsc_copy_err(char *dst, long dst_len, const char *src) {
    size_t n;
    if (!dst || dst_len <= 0) return;
    if (!src) {
        dst[0] = 0;
        return;
    }
    n = (size_t)(dst_len - 1);
    strncpy(dst, src, n);
    dst[n] = 0;
}

long wsc_open(const char *url,
              long timeout_ms,
              long verify_peer,
              long verify_host,
              void **out_handle,
              char *errbuf,
              long errbuf_len) {
    CURL *curl;
    CURLcode rc;
    char local_err[CURL_ERROR_SIZE];

    if (out_handle) *out_handle = NULL;
    local_err[0] = 0;

    curl = curl_easy_init();
    if (!curl) {
        wsc_copy_err(errbuf, errbuf_len, "curl_easy_init returned NULL");
        return CURLE_FAILED_INIT;
    }

    rc = curl_easy_setopt(curl, CURLOPT_ERRORBUFFER, local_err);
    if (rc != CURLE_OK) goto fail;

    rc = curl_easy_setopt(curl, CURLOPT_URL, url);
    if (rc != CURLE_OK) goto fail;

    rc = curl_easy_setopt(curl, CURLOPT_CONNECT_ONLY, 2L);
    if (rc != CURLE_OK) goto fail;

    rc = curl_easy_setopt(curl, CURLOPT_TIMEOUT_MS, timeout_ms);
    if (rc != CURLE_OK) goto fail;

    rc = curl_easy_setopt(curl, CURLOPT_SSL_VERIFYPEER, verify_peer ? 1L : 0L);
    if (rc != CURLE_OK) goto fail;

    rc = curl_easy_setopt(curl, CURLOPT_SSL_VERIFYHOST, verify_host ? 2L : 0L);
    if (rc != CURLE_OK) goto fail;

    rc = curl_easy_perform(curl);
    if (rc != CURLE_OK) goto fail;

    if (out_handle) *out_handle = (void *)curl;
    wsc_copy_err(errbuf, errbuf_len, "");
    return CURLE_OK;

fail:
    if (local_err[0]) wsc_copy_err(errbuf, errbuf_len, local_err);
    else wsc_copy_err(errbuf, errbuf_len, curl_easy_strerror(rc));
    curl_easy_cleanup(curl);
    return (long)rc;
}

void wsc_close(void *h) {
    if (h) curl_easy_cleanup((CURL *)h);
}

long wsc_send_text_utf8(void *h,
                        const unsigned char *buf,
                        size_t buf_len,
                        size_t *sent_bytes,
                        char *errbuf,
                        long errbuf_len) {
    CURLcode rc;
    size_t sent = 0;

    if (sent_bytes) *sent_bytes = 0;
    if (!h) {
        wsc_copy_err(errbuf, errbuf_len, "NULL handle");
        return CURLE_FAILED_INIT;
    }

    rc = curl_ws_send((CURL *)h,
                      buf,
                      buf_len,
                      &sent,
                      0,
                      CURLWS_TEXT);

    if (sent_bytes) *sent_bytes = sent;

    if (rc != CURLE_OK) {
        wsc_copy_err(errbuf, errbuf_len, curl_easy_strerror(rc));
        return (long)rc;
    }

    wsc_copy_err(errbuf, errbuf_len, "");
    return CURLE_OK;
}

long wsc_recv_text_utf8(void *h,
                        unsigned char *out_buf,
                        size_t out_buf_len,
                        size_t *received_bytes,
                        char *errbuf,
                        long errbuf_len) {
    CURLcode rc;
    size_t received = 0;
    const struct curl_ws_frame *meta = NULL;

    if (received_bytes) *received_bytes = 0;
    if (!h) {
        wsc_copy_err(errbuf, errbuf_len, "NULL handle");
        return CURLE_FAILED_INIT;
    }
    if (!out_buf || out_buf_len == 0) {
        wsc_copy_err(errbuf, errbuf_len, "output buffer missing");
        return CURLE_BAD_FUNCTION_ARGUMENT;
    }

    rc = curl_ws_recv((CURL *)h, out_buf, out_buf_len, &received, &meta);

    if (received_bytes) *received_bytes = received;

    if (rc != CURLE_OK) {
        wsc_copy_err(errbuf, errbuf_len, curl_easy_strerror(rc));
        return (long)rc;
    }

    wsc_copy_err(errbuf, errbuf_len, "");
    return CURLE_OK;
}

const char *wsc_last_error_text(long code) {
    return curl_easy_strerror((CURLcode)code);
}

#ifdef __cplusplus
}
#endif