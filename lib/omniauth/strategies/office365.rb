require 'omniauth-oauth2'

module OmniAuth
  module Strategies
    class Office365 < OmniAuth::Strategies::OAuth2
      option :name, :office365

      option :client_options, {
          site:          "https://login.microsoftonline.com",
          authorize_url: "/common/oauth2/v2.0/authorize",
          token_url:     "/common/oauth2/v2.0/token"
      }

      option :authorize_params, {
        scope: 'https://outlook.office.com/Calendars.ReadWrite openid email profile offline_access'
      }

      uid { raw_info["Id"] }

      info do
        {
          'email' => raw_info["Mail"],
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
        @raw_info ||= access_token.get("https://outlook.office.com/api/v2.0/me/").parsed
      end
    end
  end
end
