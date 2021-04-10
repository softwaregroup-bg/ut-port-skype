const hrtime = require('browser-process-hrtime');
const request = (process.type === 'renderer') ? require('ut-browser-request') : require('request');
const jwt = require('jsonwebtoken');
const jwksRsa = require('jwks-rsa');
const {PassThrough} = require('readable-stream');
const {URL} = require('url');
const mailToRegEx = /<a href="mailto:([^"]+)">.+?<\/a>/g;
const sanitize = text => (typeof text === 'string') ? text.replace(mailToRegEx, '$1') : text; // replace all <a href="mailto:...">...</a>
module.exports = function skype({utBus, utMethod}) {
    const tokens = {};
    const getToken = ({appId, secret}) => {
        let token = tokens[appId];
        if (token && hrtime(token.time)[0] < token.expires_in - 30) {
            return token.access_token;
        } else {
            return new Promise((resolve, reject) => request.post({
                url: 'https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token',
                form: {
                    grant_type: 'client_credentials',
                    client_id: appId,
                    client_secret: secret,
                    scope: 'https://api.botframework.com/.default'
                }
            }, (error, response, body) => {
                if (error || response.statusCode < 200 || response.statusCode >= 300) {
                    reject(new Error('Bot authentication error'));
                } else {
                    token = JSON.parse(response.body);
                    token.time = hrtime();
                    tokens[appId] = token;
                    resolve(token.access_token);
                }
            }));
        }
    };
    const image = attachments => attachments
        .filter(image => typeof image === 'string' || /^image\/(jpeg|png|gif)$/.test(image.contentType))
        .map(image => typeof image === 'string' ? {
            name: 'image',
            contentUrl: image,
            thumbnailUrl: image
        } : {
            contentType: image.contentType,
            name: image.title,
            contentUrl: image.url,
            thumbnailUrl: image.thumbnail
        });
    const quickButtons = attachments => attachments
        .filter(button => typeof button === 'string' || button.contentType === 'application/x.button')
        .map(button => typeof button === 'string' ? {
            type: 'imBack',
            title: button,
            value: button
        } : {
            type: 'imBack',
            title: button.title || button.value,
            value: button.value
        });

    const type = button => ({
        url: 'openUrl',
        reply: 'imBack',
        post: 'postBack'
    }[button.details && button.details.type] || 'openUrl');

    const richButtons = attachments => attachments
        .filter(button => typeof button === 'string' || button.contentType === 'application/x.button')
        .map(button => typeof button === 'string' ? {
            type: 'imBack',
            title: button,
            value: button
        } : {
            type: type(button),
            title: button.title,
            value: type(button) === 'openUrl' ? button.url : button.value,
            image: button.thumbnail
        });
    const locationEntities = attachments => attachments
        .filter(location => location.contentType === 'application/x.location' && location.details)
        .map(location => ({
            type: 'Place',
            geo: {
                type: 'GeoCoordinates',
                name: location.title,
                latitude: location.details.lat,
                longitude: location.details.lon
            }
        }));
    const location = attachments => attachments
        .filter(location => location.contentType === 'application/x.location' && location.details)
        .map(location => ({
            contentType: 'application/vnd.microsoft.card.hero',
            name: 'map',
            content: {
                title: location.title,
                text: location.details.address,
                images: [{
                    url: location.thumbnail
                }],
                buttons: [{
                    type: 'openUrl',
                    title: 'open',
                    value: location.url
                }]
            },
            thumbnailUrl: location.thumbnail
        }));
    const list = attachments => attachments
        .filter(button => typeof button === 'string' || button.contentType === 'application/x.button')
        .map(button => ({
            contentType: 'application/vnd.microsoft.card.hero',
            // name: 'list',
            content: {
                title: button.title,
                text: button.details && button.details.subtitle,
                images: button.thumbnail && [{
                    url: button.thumbnail
                }],
                tap: (button.url || button.value) && {
                    type: button.url ? 'openUrl' : 'postBack',
                    title: button.title,
                    value: button.url || button.value
                },
                buttons: button.details && Array.isArray(button.details.actions) && button.details.actions.map(action => ({
                    type: 'openUrl',
                    title: action.title,
                    value: action.url
                }))
            }
        }));
    const sendMessage = async(msg, {auth}) => {
        if (!msg) return msg;
        const from = {id: msg.sender.id};
        const recipient = {id: msg.receiver.id};
        const url = msg.receiver.conversationId;
        const headers = {
            Authorization: 'Bearer ' + await getToken(auth)
        };
        switch (msg && msg.type) {
            case 'text': return {
                url,
                headers,
                body: {
                    from,
                    recipient,
                    type: 'message',
                    text: msg.text
                }
            };
            case 'location': return {
                url,
                headers,
                body: {
                    from,
                    recipient,
                    type: 'message',
                    text: msg.text,
                    attachments: location(msg.attachments),
                    entities: locationEntities(msg.attachments)
                }
            };
            case 'image': return {
                url,
                headers,
                body: {
                    from,
                    recipient,
                    type: 'message',
                    text: msg.text,
                    attachments: image(msg.attachments)
                }
            };
            case 'quick': return {
                url,
                headers,
                body: {
                    from,
                    recipient,
                    type: 'message',
                    attachments: [{
                        contentType: 'application/vnd.microsoft.card.hero',
                        name: 'quick answers',
                        content: {
                            title: msg.details && msg.details.title,
                            text: msg.text,
                            buttons: quickButtons(msg.attachments)
                        },
                        thumbnailUrl: msg.thumbnail
                    }]
                }
            };
            case 'actions': return {
                url,
                headers,
                body: {
                    from,
                    recipient,
                    type: 'message',
                    attachments: [{
                        contentType: 'application/vnd.microsoft.card.hero',
                        name: 'actions',
                        content: {
                            title: msg.details && msg.details.title,
                            text: msg.text,
                            buttons: richButtons(msg.attachments)
                        },
                        thumbnailUrl: msg.thumbnail
                    }]
                }
            };
            case 'list': return {
                url,
                headers,
                body: {
                    from,
                    recipient,
                    type: 'message',
                    attachmentLayout: 'carousel',
                    attachments: list(msg.attachments)
                }
            };
            default: return false;
        }
    };
    return class skype extends require('ut-port-webhook')(...arguments) {
        get defaults() {
            return {
                path: '/skype/{appId}/{clientId}',
                hook: 'skypeIn',
                namespace: 'skype',
                server: {
                    port: 8080
                },
                openIdUrl: 'https://login.botframework.com/v1/.well-known/openidconfiguration'
            };
        }

        async start(...params) {
            utBus.attachHandlers(this.methods, [this.config.id.replace('skype', 'webchat')]);
            this.httpServer.route({
                method: 'GET',
                path: this.config.path + '/{channelId}/attachment',
                options: {
                    auth: false,
                    handler: async({query, params}, h) => {
                        const {url} = query;
                        if (!url) return h.response().code(404);
                        const auth = await utMethod('bot.botContext.fetch#[0]')({
                            platform: 'skype',
                            appId: params.appId,
                            clientId: params.channelId + '/' + params.clientId
                        });
                        return h.response(request({
                            url,
                            headers: {
                                Authorization: 'Bearer ' + await getToken(auth)
                            }
                        }).pipe(new PassThrough()));
                    }
                }
            });
            return super.start(...params);
        }

        handlers() {
            let openIdMetaDoc;
            let jwksClient;
            const getKeyFromJwtHeader = ({kid}, callback) => {
                jwksClient.getSigningKey(kid, (error, key) => {
                    if (error) {
                        callback(error);
                    } else {
                        callback(null, key.publicKey || key.rsaPublicKey);
                    }
                });
            };
            const {namespace, hook} = this.config;
            return {
                async ready() {
                    openIdMetaDoc = await this.sendRequest({
                        method: 'GET',
                        url: this.config.openIdUrl
                    });
                    jwksClient = jwksRsa({ // TODO add caching options
                        jwksUri: openIdMetaDoc.jwks_uri
                    });
                },
                [`${hook}.identity.request.receive`]: async(msg, {params, request: {headers}}) => {
                    // https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-authentication?view=azure-bot-service-4.0#connector-to-bot
                    if (typeof headers.authorization !== 'string') {
                        throw this.errors['webhook.missingHeader']({params: {header: 'authorization'}});
                    }
                    const [scheme, token] = headers.authorization.split(' ');
                    if (scheme !== 'Bearer' || !token) {
                        throw this.errors['webhook.malformedHeader']({params: {header: 'authorization'}});
                    }
                    let error;
                    let decoded;
                    try {
                        decoded = await new Promise((resolve, reject) => {
                            return jwt.verify(
                                token,
                                getKeyFromJwtHeader,
                                {
                                    algorithms: openIdMetaDoc.id_token_signing_alg_values_supported,
                                    clockTolerance: 5 * 60
                                },
                                (err, res) => err ? reject(err) : resolve(res)
                            );
                        });
                    } catch (e) {
                        error = e;
                    }
                    if (
                        error ||
                        decoded.aud !== params.appId ||
                        decoded.iss !== 'https://api.botframework.com' ||
                        decoded.serviceurl !== msg.serviceUrl
                    ) {
                        throw this.errors['webhook.integrityValidationFailed'](error);
                    }

                    return {
                        platform: 'skype',
                        appId: params.appId,
                        clientId: msg.channelId + '/' + params.clientId
                    };
                },
                [`${hook}.identity.response.send`]: msg => {
                    return msg;
                },
                [`${hook}.message.request.receive`]: (msg, $meta) => {
                    const event = msg.type === 'event' && msg.name;
                    const text = event ? msg.value : sanitize(msg.text);
                    return {
                        type: 'text',
                        messageId: msg.id,
                        timestamp: new Date(msg.timestamp).getTime(),
                        sender: {
                            id: msg.from.id,
                            platform: 'skype',
                            conversationId: new URL(`v3/conversations/${msg.conversation.id}/activities`, msg.serviceUrl).href,
                            // threadId: new URL(`v3/conversations/${msg.conversation.id}/activities/${msg.id}`, msg.serviceUrl).href,
                            contextId: $meta.auth.contextId
                        },
                        receiver: {
                            id: msg.recipient.id
                        },
                        event,
                        text,
                        request: msg,
                        attachments: msg.attachments && msg.attachments.map(({contentType, contentUrl, name}) => {
                            const download = new URL(`${$meta.url.pathname}/${msg.channelId}/attachment`, this.getUriFromMeta($meta));
                            download.searchParams.set('url', contentUrl);
                            return {
                                url: download.href,
                                contentType,
                                filename: name
                            };
                        })
                    };
                },
                [`${namespace}.message.send.request.send`]: sendMessage,
                [`${namespace}.message.send.response.receive`]: () => false,
                [`${hook}.message.response.send`]: sendMessage
            };
        }
    };
};
