const common = joi => joi.object({
    url: /^.+\/v3\/conversations\/[^\\]+\/activities\/[^\\]+$/,
    headers: joi.object({
        Authorization: joi.string().regex(/^Bearer .+/)
    })
})
    .meta({
        apiDoc: 'https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-quickstart?view=azure-bot-service-4.0#reply-to-the-users-message'
    });

module.exports = joi => joi.object({
    body: joi.object({
        type: 'message',
        from: joi.object({
            id: joi.string().required(),
            name: joi.string().required()
        }),
        recipient: joi.object({
            id: joi.string().required(),
            name: joi.string().required()
        }),
        text: joi.string().required(),
        entities: joi.array(),
        attachments: joi.array().items(joi.object({
            contentType: joi.string().allow([
                'image/png',
                'image/jpeg',
                'application/vnd.microsoft.card.hero'
            ]),
            name: joi.string(),
            content: joi.object({
                buttons: joi.array().items(joi.object({
                    type: joi.string().allow([
                        'imBack'
                    ]).required(),
                    value: joi.string().required(),
                    image: joi.string(),
                    title: joi.string().required(),
                    text: joi.string().required()
                })),
                title: joi.string.required()
            }),
            contentUrl: joi.string().uri({scheme: ['http', 'https']}),
            thumbnailUrl: joi.string().uri({scheme: ['http', 'https']}).required()
        }).xor('content', 'contentUrl'))
            .meta({
                apiDoc: 'https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-api-reference?view=azure-bot-service-4.0#attachment-object'
            })
            .description('Defines additional information to include in the message. An attachment may be a media file (e.g., audio, video, image, file) or a rich card.')
    })
})
    .concat(common(joi))
    .meta({
        apiDoc: 'https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-api-reference?view=azure-bot-service-4.0#reply-to-activity'
    })
    .description('Message sent to Skype');
