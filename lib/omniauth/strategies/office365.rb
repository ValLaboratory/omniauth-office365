require 'omniauth-oauth2'

module OmniAuth
  module Strategies
    class Office365 < OmniAuth::Strategies::OAuth2
      option :name, :office365

      option :client_options, {
        site:          'https://outlook.office.com/api/v2.0',
        token_url:     'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        authorize_url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
      }

      option :authorize_params, {
        scope: 'https://outlook.office.com/Calendars.ReadWrite offline_access'
      }

      uid { raw_info["MailboxGuid"] }

      info do
        {
          'email' => raw_info["Id"],
          'name' => raw_info["DisplayName"],
          'nickname' => raw_info["Alias"]
        }
      end

      extra do
        {
          'raw_info' => raw_info
        }
      end

      def raw_info
        @raw_info ||= access_token.get('https://outlook.office.com/api/v1.0/users/me').parsed
      end
    end
  end
end
